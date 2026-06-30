import os
import re
from collections import deque
from pathlib import Path
from date_extraction import extract_year_key


def _calculate_closeness_scores(text_content, search_words):
    if not search_words:
        return []
        
    search_words_set = set(w.lower() for w in search_words)
    if len(search_words_set) < 2:
        return []
    
    # OPTIMIZATION: Use finditer to avoid loading massive list of strings into memory 
    # at once, and only lowercase the matched word, not the entire text_content.
    occurrences = []
    for i, match in enumerate(re.finditer(r'\b\w+\b', text_content)):
        word = match.group().lower()
        if word in search_words_set:
            occurrences.append({'idx': i, 'word': word})

    # Early exit if we didn't find at least one of every required search word
    if not occurrences or len(set(occ['word'] for occ in occurrences)) < len(search_words_set):
        return []

    scores = []
    window = deque()
    word_counts_in_window = {}
    
    for right_item in occurrences:
        window.append(right_item)
        word_counts_in_window[right_item['word']] = word_counts_in_window.get(right_item['word'], 0) + 1

        while len(word_counts_in_window) == len(search_words_set):
            start_idx = window[0]['idx']
            end_idx = window[-1]['idx']
            
            span_length = end_idx - start_idx + 1
            score = span_length - len(search_words_set)
            scores.append(score)

            left_item = window.popleft()
            word_counts_in_window[left_item['word']] -= 1
            if word_counts_in_window[left_item['word']] == 0:
                del word_counts_in_window[left_item['word']]

    return sorted(scores)


def _create_search_pattern(part):
    val = part["value"]
    if part["type"] == "phrase":
        return re.compile(re.escape(val), re.IGNORECASE)
    if val.endswith('*'):
        base = val[:-1]
        return re.compile(rf'\b{re.escape(base)}\w*\b', re.IGNORECASE)
    elif val.endswith('?'):
        base = val[:-1]
        return re.compile(rf'\b{re.escape(base)}\w\b', re.IGNORECASE)
    elif val.isdigit():
        return re.compile(re.escape(val), re.IGNORECASE)
    else:
        return re.compile(rf'\b{re.escape(val)}(s)?\b', re.IGNORECASE)


def _get_item_tags(app, dir_tree, path, is_folder=False):
    if not app:
        return ()

    path_str = str(path)
    # Using generator is fine, but string startswith is fast
    parent_dir = next((d for d in app.directories if path_str.startswith(d)), None)
    
    if parent_dir:
        # OPTIMIZATION: Replaced Path(path_str).relative_to() with raw string manipulation. 
        # Path() object instantiation is slow, string manipulation + os.sep is extremely fast.
        rel_str = path_str[len(parent_dir):].strip(os.sep)
        parts_len = len(rel_str.split(os.sep)) if rel_str else 0
        depth = parts_len if is_folder else max(0, parts_len - 1)
    else:
        depth = 0

    base_color = app.directory_colors.get(parent_dir, "#ffffff")
    row_bg = app._get_depth_color(base_color, depth)
    txt_color = app._get_contrast_color(row_bg)
    tag_name = f"bg_{row_bg}"
    
    dir_tree.tag_configure(tag_name, background=row_bg, foreground=txt_color)
    return (tag_name,)


def execute_search(query, cache, directory_active_status, dir_tree, on_view_mode_change_callback, app=None):
    if not query:
        on_view_mode_change_callback()
        return

    # --- Setup and Parse Search Query ---
    year_sort = False
    q_lower = query.lower()
    if q_lower.endswith(" ayo"):
        year_sort = True
        query = query[:-len(" ayo")].rstrip()
    elif q_lower == "ayo":
        year_sort = True
        query = ""

    search_parts = []
    if query:
        for match in re.finditer(r'\(([^)]+)\)|"([^"]*)"|(\S+)', query):
            if match.group(1): # Path only
                search_parts.append({"type": "phrase", "value": match.group(1).strip(), "scope": "path"})
            elif match.group(2): # Exact Phrase
                search_parts.append({"type": "phrase", "value": match.group(2).strip(), "scope": "all"})
            else: # Standard Word
                search_parts.append({"type": "word", "value": match.group(3).strip(), "scope": "all"})

    if not search_parts and not year_sort:
        on_view_mode_change_callback()
        return

    # OPTIMIZATION: Pre-compile regular expressions BEFORE entering massive nested loops
    for part in search_parts:
        part["pattern"] = _create_search_pattern(part)

    # --- AYO (year sorting) logic ---
    if year_sort:
        matches_with_counts = []
        for directory, data in cache.items():
            if not directory_active_status.get(directory, True):
                continue
                
            for path, text in zip(data["files"], data["texts"]):
                if not search_parts:
                     matches_with_counts.append((path, 0))
                     continue

                # OPTIMIZATION: Removed expensive .lower() conversions. Regexes are already compiled with re.IGNORECASE.
                str_path = str(path)
                all_parts_found = True
                
                for part in search_parts:
                    pat = part["pattern"]  # Use Pre-compiled regex
                    
                    if part["scope"] == "path":
                        if not pat.search(str_path):
                            all_parts_found = False
                            break
                    else:
                        if not (pat.search(text) or pat.search(str_path)):
                            all_parts_found = False
                            break
                
                if all_parts_found:
                    matches_with_counts.append((path, 1))

        matches_with_counts.sort(key=lambda x: extract_year_key(x[0]), reverse=True)
        dir_tree.delete(*dir_tree.get_children())
        
        for path, count in matches_with_counts:
            display = f"{os.path.basename(path)}" if not search_parts else f"{os.path.basename(path)} ({count})"
            tags = _get_item_tags(app, dir_tree, path, is_folder=False)
            dir_tree.insert("", "end", text=display, values=(path,), tags=tags)
        return
    # --- End of AYO logic ---


    # --- Standard Search Logic ---
    search_words_for_closeness = list(set(
        part['value'].lower() for part in search_parts if part['type'] == 'word' and part['scope'] == 'all'
    ))

    file_matches = []
    folder_matches = set()
    all_subfolders = set()
    
    # --- Collect all active subfolders and search files ---
    for directory, data in cache.items():
        if not directory_active_status.get(directory, True):
            continue

        all_subfolders.add(str(directory))

        for path, text in zip(data["files"], data["texts"]):
            str_path = str(path) # Calculate string once per loop iteration
            all_subfolders.add(os.path.dirname(str_path))
            
            all_parts_found_in_file = True
            total_occurrences = 0
            
            for part in search_parts:
                pat = part["pattern"] # Use Pre-compiled regex
                path_occ = len(pat.findall(str_path))
                
                if part["scope"] == "path":
                    if path_occ > 0:
                        total_occurrences += path_occ
                    else:
                        all_parts_found_in_file = False
                        break
                else:
                    text_occ = len(pat.findall(text))
                    if text_occ > 0 or path_occ > 0:
                        total_occurrences += (text_occ + path_occ)
                    else:
                        all_parts_found_in_file = False
                        break
            
            if all_parts_found_in_file:
                closeness_scores = _calculate_closeness_scores(text, search_words_for_closeness)
                file_matches.append({
                    "path": path,
                    "count": total_occurrences,
                    "scores": closeness_scores if closeness_scores else [float('inf')]
                })

    # --- Search Collected Folders ---
    for folder_path in all_subfolders:
        folder_name = os.path.basename(os.path.normpath(folder_path))
        if not folder_name:
            folder_name = str(folder_path)
            
        folder_is_match = True
        for part in search_parts:
            # OPTIMIZATION: Pre-compiled IGNORECASE regex skips string lowercase operations
            if not part["pattern"].search(folder_name):
                folder_is_match = False
                break
                
        if folder_is_match:
            folder_matches.add(folder_path)

    # --- Sort Results ---
    file_matches.sort(key=lambda x: (x["scores"], -x["count"]))
    sorted_folder_matches = sorted(list(folder_matches))

    # --- Update TreeView ---
    dir_tree.delete(*dir_tree.get_children())

    for path in sorted_folder_matches:
        display = f"📁 {os.path.basename(os.path.normpath(path))}"
        tags = _get_item_tags(app, dir_tree, path, is_folder=True)
        dir_tree.insert("", "end", text=display, values=(path,), tags=tags)

    for match in file_matches:
        display = f"{os.path.basename(match['path'])} ({match['count']})"
        tags = _get_item_tags(app, dir_tree, match['path'], is_folder=False)
        dir_tree.insert("", "end", text=display, values=(match['path'],), tags=tags)