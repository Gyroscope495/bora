import os
import time
from pathlib import Path
from collections import defaultdict
import documentinfo

def get_next_view_mode(current_mode):
    """Cycles through the view modes and returns the new mode and its color."""
    if current_mode == "Cached tree":
        return "Saved", "#cce5ff"
    elif current_mode == "Saved":
        return "Recent", "#e6e6fa"
    else:  # Was "Recent"
        return "Cached tree", "#ccffcc"

def get_color_tag_for_path(app, path_str):
    """Finds the base directory for a path, calculates its depth, and returns the appropriate color tag."""
    path_obj = Path(path_str).resolve()
    best_match = None
    
    # Find the most specific root directory this path belongs to
    for d in app.directories:
        try:
            # Check if path_obj is inside Path(d)
            path_obj.relative_to(Path(d).resolve())
            # Keep the longest matching directory path
            if best_match is None or len(str(Path(d))) > len(str(Path(best_match))):
                best_match = d
        except ValueError:
            continue
            
    if best_match:
        base_color = app.directory_colors.get(best_match, "#ffffff")
        try:
            # Calculate how many levels deep the file/folder is from its root
            depth = len(path_obj.relative_to(Path(best_match).resolve()).parts)
        except Exception:
            depth = 0
            
        # Tap into the app's cascading color and contrast math
        node_color = app._get_depth_color(base_color, depth)
        fg_color = app._get_contrast_color(node_color)
        tag_name = f"bg_{node_color}"
        
        # Configure the tag in the Treeview
        app.dir_tree.tag_configure(tag_name, background=node_color, foreground=fg_color)
        return tag_name
        
    return ""

def build_saved_files_tree(app):
    """Populates the tree with saved files, organized hierarchically with colors."""
    inserted_nodes = {}
    sorted_saved_files = sorted(list(app.saved_files))
    
    for file_path in sorted_saved_files:
        path_obj = Path(file_path).resolve()
        # Get all parent folders from the drive root down to the file's folder
        parents = list(path_obj.parents)[::-1] 
        
        parent_iid = ""
        for p in parents:
            p_str = str(p)
            if p_str not in inserted_nodes:
                text = p.name if p.name else p_str
                tag = get_color_tag_for_path(app, p_str)
                
                # Insert folder node with open=True and apply color tag if it exists
                if tag:
                    parent_iid = app.dir_tree.insert(parent_iid, "end", text=text, values=(p_str,), open=True, tags=(tag,))
                else:
                    parent_iid = app.dir_tree.insert(parent_iid, "end", text=text, values=(p_str,), open=True)
                inserted_nodes[p_str] = parent_iid
            else:
                parent_iid = inserted_nodes[p_str]
                
        # Insert the actual file with its corresponding color tag
        tag = get_color_tag_for_path(app, str(path_obj))
        if tag:
            app.dir_tree.insert(parent_iid, "end", text=path_obj.name, values=(str(path_obj),), tags=(tag,))
        else:
            app.dir_tree.insert(parent_iid, "end", text=path_obj.name, values=(str(path_obj),))

def generate_recent_files_mindmap(recent_files, main_directories):
    """
    Generates a mindmap string from a list of recent files,
    grouped by one folder level below the main directories.
    This version uses a simple indented list that works with proportional fonts.
    """
    if not recent_files:
        return "Recent Files Mindmap\n\nNo recent files found in the specified timespan."

    # Helper to find which main directory a path belongs to.
    sorted_main_dirs = sorted(main_directories, key=len, reverse=True)
    def find_base_dir(path):
        for base in sorted_main_dirs:
            try:
                path.relative_to(base)
                return base
            except ValueError:
                continue
        return None

    # Group files by subdirectory. Files in each group maintain their recency order.
    files_by_group_dir = defaultdict(list)
    for _, file_path in recent_files:
        p = Path(file_path)
        base_dir = find_base_dir(p)
        if not base_dir: continue

        group_dir_path = base_dir
        try:
            relative_path = p.relative_to(base_dir)
            if len(relative_path.parts) > 1:
                group_dir_path = str(Path(base_dir) / relative_path.parts[0])
            else:
                group_dir_path = str(base_dir)
        except ValueError:
             group_dir_path = str(base_dir)
        files_by_group_dir[group_dir_path].append(p.name)

    # Build the mindmap string using a simple list format.
    mindmap_lines = ["Recent Files Mindmap"]
    sorted_group_dirs = sorted(files_by_group_dir.keys())

    for directory in sorted_group_dirs:
        # Add a blank line before each folder group for better separation.
        mindmap_lines.append("")
        mindmap_lines.append(f"📁 {os.path.basename(directory)}")
        
        files_in_dir = files_by_group_dir[directory]
        for filename in files_in_dir:
            mindmap_lines.append(f"  - {filename}")

    return "\n".join(mindmap_lines)


def build_recent_files_tree(app):
    """
    Populates the tree with recent files organized hierarchically with colors.
    Returns a list of recent files sorted by modification time.
    """
    now = time.time()
    timespan_seconds = app.recent_timespan_hours * 3600
    recent_files = []

    for dir_path, dir_data in app.cache.items():
        if not app.directory_active_status.get(dir_path, True):
            continue
        
        for file_path in dir_data.get("files", []):
            try:
                mtime = os.path.getmtime(file_path)
                if (now - mtime) <= timespan_seconds:
                    recent_files.append((mtime, file_path))
            except (OSError, FileNotFoundError):
                continue
    
    # Sort by modification time (mtime), descending (most recent first)
    sorted_recent = sorted(recent_files, key=lambda x: x[0], reverse=True)
    
    # Dictionary to keep track of folder nodes we've already created
    inserted_nodes = {}

    for mtime, file_path in sorted_recent:
        path_obj = Path(file_path).resolve()
        
        # Build the parent folder structure for this file
        parents = list(path_obj.parents)[::-1]
        parent_iid = ""
        
        for p in parents:
            p_str = str(p)
            if p_str not in inserted_nodes:
                text = p.name if p.name else p_str
                tag = get_color_tag_for_path(app, p_str)
                
                # Insert folder node with open=True and apply color tag if it exists
                if tag:
                    parent_iid = app.dir_tree.insert(parent_iid, "end", text=text, values=(p_str,), open=True, tags=(tag,))
                else:
                    parent_iid = app.dir_tree.insert(parent_iid, "end", text=text, values=(p_str,), open=True)
                inserted_nodes[p_str] = parent_iid
            else:
                parent_iid = inserted_nodes[p_str]
                
        # Insert the actual file with its corresponding color tag
        tag = get_color_tag_for_path(app, str(path_obj))
        file_name = path_obj.name
        
        if tag:
            app.dir_tree.insert(parent_iid, "end", text=file_name, values=(str(path_obj),), tags=(tag,))
        else:
            app.dir_tree.insert(parent_iid, "end", text=file_name, values=(str(path_obj),))
            
    return sorted_recent

def on_view_mode_change(app, *_):
    """Clears and rebuilds the directory tree based on the current view mode."""
    mode = app.view_mode.get()
    
    # Clear existing tree items
    for iid in app.dir_tree.get_children():
        app.dir_tree.delete(iid)

    # Rebuild based on mode and update the info panel accordingly
    if mode == "Cached tree":
        documentinfo.clear_doc_info(app.output_text, app.published_var)
        for directory in sorted(app.directories):
            app.build_directory_tree(directory)
    elif mode == "Saved":
        documentinfo.clear_doc_info(app.output_text, app.published_var)
        build_saved_files_tree(app)
    elif mode == "Recent":
        recent_files_list = build_recent_files_tree(app)
        # Pass app.directories to the mindmap generator for grouping context
        mindmap_text = generate_recent_files_mindmap(recent_files_list, app.directories)
        documentinfo.display_summary_text(app.output_text, mindmap_text)
    
    # Update the tree heading
    app.dir_tree.heading("#0", text=mode)