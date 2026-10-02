"""Shared helpers used by the COPQ / FTT pipeline scripts."""

import os
import re

# Site codes, ordered so longer codes (VH2) are matched before shorter (VH).
SITE_CODE_CANDIDATES = ["VH4", "VH3", "VH2", "VH", "JV3", "JV2", "JVB", "JV"]


def extract_site_name(filename):
    """Extract a standard site name from a filename (VH, VH2, VH3, VH4,
    JV, JV2, JV3). JVB is normalized to JV3 to stay consistent with the
    site's actual name used elsewhere in the pipeline."""
    base_name = os.path.splitext(os.path.basename(filename))[0].upper().strip()

    def normalize(code):
        return "JV3" if code == "JVB" else code

    # 1) Prefer a token that is *exactly* a site code once split on
    #    common separators - the most reliable signal.
    for tok in re.split(r"[-_ ]+", base_name):
        if tok in SITE_CODE_CANDIDATES:
            return normalize(tok)

    # 2) Substring search bounded by non-alphanumeric characters, for
    #    filenames without separators (e.g. "VH2report.xlsx").
    for cand in SITE_CODE_CANDIDATES:
        match = re.search(rf"(?<![A-Za-z0-9]){cand}(?![A-Za-z0-9])", base_name)
        if match:
            return normalize(match.group(0))

    # 3) Last resort: trailing "word" / dash-segment heuristic.
    words = base_name.split()
    last_word = words[-1] if words else base_name
    candidate = last_word.split("-")[-1] if "-" in last_word else last_word
    return candidate


def save_with_fallback(book, output_path, max_retries=19):
    """Save `book` (an openpyxl Workbook or pptx Presentation) to
    output_path, falling back to suffixed names (_v1, _v2, ...) if the
    original file is locked (e.g. open in Excel/PowerPoint).

    Returns the path actually written.
    """
    base, ext = os.path.splitext(output_path)
    try:
        book.save(output_path)
        return output_path
    except (PermissionError, OSError) as e:
        print(f"\n[!] WARNING: Could not save to '{output_path}' ({e}). Attempting fallback names...")
        for i in range(1, max_retries + 1):
            fallback_file = f"{base}_v{i}{ext}"
            try:
                book.save(fallback_file)
                print(f"[+] Report saved to fallback file instead: {fallback_file}")
                return fallback_file
            except (PermissionError, OSError):
                continue
        print(f"[!] ERROR: Failed to save '{output_path}' after {max_retries} fallback attempts.")
        raise
