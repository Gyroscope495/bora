# zoom_view.py
"""
This module provides the "Zoom" view functionality for the Bora application.
It sorts files based on their date-like prefixes relative to a selected file.

Update (search-aware):
- When the search field has text, Zoom lists ONLY files that also match the
  existing search logic (content + filename), reusing the app's cache.
- If the search text is only the special token "ayo" (year sort toggle), no
  filtering is applied.
"""

import os
import re
from pathlib import Path
from date_extraction import parse_token_date, extract_year_key

# -------------------- Search helpers --------------------

def _get_search_text(app):
    """
    Return the current search query from the host app (stripped), or "" if none.
    Tries app.search_var (tk.StringVar) first, then app.search_entry (Entry).
    """
    q = ""
    try:
        if hasattr(app, "search_var") and app.search_var:
            q = app.search_var.get() or ""
    except Exception:
        pass
    if not q:
        try:
            if hasattr(app, "search_entry") and app.search_entry:
                q = app.search_entry.get() or ""
        except Exception:
            pass
    return q.strip()


def _compute_search_parts_and_flag(query):
    """Mirror Bora's tokenization & 'ayo' handling."""
    year_sort = False
    q_lower = query.lower()
    if q_lower.endswith(" ayo"):
        year_sort = True
        query = query[:-len(" ayo")].rstrip()
    elif q_lower == "ayo":
        year_sort = True
        query = ""

    search_parts = []
    for match in re.finditer(r'"([^"]*)"|(\S+)', query):
        if match.group(1):
            search_parts.append({"type": "phrase", "value": match.group(1)})
        else:
            search_parts.append({"type": "word", "value": match.group(2)})

    return search_parts, year_sort


def _gather_search_matches_dict(app, query):
    """
    Reuse Bora's search semantics to produce the dict of matching file paths to counts.
    Returns:
        - dict(path: count) if there are actual search tokens (after handling 'ayo')
        - None if there are no tokens (so Zoom shouldn't filter)
    """
    search_parts, year_sort = _compute_search_parts_and_flag(query)
    if not search_parts and not year_sort:
        return None  # no actual search; don't filter

    # Local import to prevent circular dependency
    from search import _create_search_pattern

    matches = {}
    # Iterate over cached texts the same way the main search does.
    for directory, data in getattr(app, "cache", {}).items():
        if not app.directory_active_status.get(directory, True):
            continue

        files = data.get("files", [])
        texts = data.get("texts", [])
        for path, text in zip(files, texts):
            total_occurrences = 0
            all_ok = True

            for part in search_parts:
                pat = _create_search_pattern(part)
                
                text_occ = len(pat.findall(text))
                path_occ = len(pat.findall(str(path)))
                
                if text_occ > 0 or path_occ > 0:
                    total_occurrences += (text_occ + path_occ)
                else:
                    all_ok = False
                    break

            if all_ok:
                matches[path] = total_occurrences

    # If only 'ayo' was present (year_sort True, but search_parts empty),
    # treat as NO filter.
    if not matches and year_sort and not search_parts:
        return None

    return matches


# -------------------- Main Zoom rendering --------------------

def apply_zoom_view(app):
    """
    Re-populates the directory tree with a "zoomed" view, sorting all active
    files chronologically relative to the currently selected file.
    If the search bar has text, only files whose contents/filenames match
    the active search are shown.
    """
    # Local import to prevent circular dependency
    from search import _get_item_tags

    focused = app.dir_tree.focus()
    target_path = app._node_path(focused) if focused else None

    base_y, base_m, base_d = (None, None, None)
    if target_path:
        base_name = os.path.splitext(os.path.basename(target_path))[0]
        tok = base_name.split()[0]
        parsed = parse_token_date(tok)
        if parsed:
            base_y, base_m, base_d = parsed

    # Search-aware filtering: intersect with current search results, if any.
    search_text = _get_search_text(app)
    search_matches = _gather_search_matches_dict(app, search_text)

    all_files = []
    filename_pattern = re.compile(r"^(?:\d{4,}|-\d{4})")
    for directory, is_active in app.directory_active_status.items():
        if not is_active:
            continue
        for root, _, filenames in os.walk(directory):
            for fname in filenames:
                if filename_pattern.match(os.path.basename(fname)):
                    file_path = os.path.join(root, fname)
                    if search_matches is not None and file_path not in search_matches:
                        continue
                    all_files.append(file_path)

    all_files.sort(key=extract_year_key, reverse=True)

    app.dir_tree.delete(*app.dir_tree.get_children())
    path_to_iid = {}
    
    search_parts, _ = _compute_search_parts_and_flag(search_text)

    for path in all_files:
        name = os.path.basename(path)
        prefix = ""
        if base_y is not None:
            base = (base_y, base_m, base_d)
            bare = os.path.splitext(name)[0]
            tok = bare.split()[0]
            dvals = parse_token_date(tok)
            if dvals:
                target = dvals
                
                # Determine direction: is the file older than our base file?
                is_past = target < base
                
                # Calculate absolute difference to avoid weird negative borrow bugs
                t1, t2 = (target, base) if is_past else (base, target)
                
                yrs = t2[0] - t1[0]
                mos = t2[1] - t1[1]
                days = t2[2] - t1[2]
                
                if days < 0:
                    mos -= 1
                if mos < 0:
                    yrs -= 1
                    mos += 12
                    
                # Prepend the minus sign if the file is in the past
                sign = "-" if is_past else ""
                prefix = f"{sign}{yrs},{mos} "

        # --- Data (Count) Resolution ---
        count_suffix = ""
        if search_matches is not None and search_parts:
            if path in search_matches:
                count_suffix = f" ({search_matches[path]})"

        # --- Color Resolution ---
        tags = _get_item_tags(app, app.dir_tree, path, is_folder=False)
        # -----------------------------

        # Insert with the specific color tag
        iid = app.dir_tree.insert("", "end", text=prefix + name + count_suffix, values=(path,), tags=tags)
        path_to_iid[path] = iid

    if target_path in path_to_iid:
        iid = path_to_iid[target_path]
        app.dir_tree.selection_set(iid)
        app.dir_tree.focus(iid)
        app.dir_tree.see(iid)

    app.view_mode.set("Zoomed")
    app.view_mode_button.config(text="Zoomed", bg="#add8e6")