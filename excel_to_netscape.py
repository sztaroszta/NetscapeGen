"""
NetscapeGen

Description:
    A cross-platform utility (macOS, Windows, Linux) to convert Excel spreadsheets
    (.xlsx) into standard Netscape-format HTML bookmark files, compatible with
    all major browsers (Chrome, Safari, Firefox, Edge).

    The program executes the following numbered steps:
      1. Launches a hidden root window to manage the GUI lifecycle.
      2. Prompts the user to define the header row configuration (Row 1 vs Row 2).
      3. Opens a native file dialog for the user to select the input Excel file.
      4. Calculates default output paths and prompts the user to confirm the
         save location for the generated HTML file.
      5. Initializes a 'ProgressLoader' to provide visual feedback via a popup
         window and simultaneous terminal output.
      6. Reads the Excel file and validates the presence of required columns
         ('Title', 'URL').
      7. Dynamically detects folder columns (e.g., 'FolderL1', 'FolderL2') using
         regex to support arbitrary directory depth.
      8. Cleans the dataset, handling missing values and converting types.
      9. Transforms the flat Excel data into a nested dictionary tree structure
         representing the bookmark hierarchy.
     10. Recursively generates Netscape-compliant HTML tags from the tree.
     11. Writes the final HTML file to disk and reveals it in the system's
         file manager (Finder, Explorer, etc.).
     12. Analyzes the generated tree to compile statistics (total bookmarks,
         folders per level).
     13. Displays a final summary window detailing the conversion results.

Usage:
    - Ensure required libraries are installed:
          pip install pandas openpyxl
    - Run the script from a terminal:
          python excel_to_netscape.py

Author:     Vitalii Starosta
GitHub:     https://github.com/sztaroszta
License:    GNU Affero General Public License v3 (AGPLv3)
"""

import collections
import html
import os
import re
import subprocess
import sys
import time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk, _tkinter

import pandas as pd

# Configuration Constants
TITLE_COLUMN = 'Title'
URL_COLUMN = 'URL'


class ProgressLoader:
    """
    A helper class that manages a Toplevel popup window containing a progress bar
    and status text. It mirrors all status updates to the system terminal.
    """

    def __init__(self, parent_root, title="Processing..."):
        """
        Initialize the progress window.

        Args:
            parent_root (tk.Tk): The root Tkinter window.
            title (str): The title to display on the progress window.
        """
        self.parent = parent_root
        self.window = tk.Toplevel(parent_root)
        self.window.title(title)

        # Window Setup
        self.window.geometry("400x150")
        self.window.attributes('-topmost', True)
        self.window.transient(parent_root)
        self.window.resizable(False, False)

        # GUI Elements
        self.center_window()
        self.main_frame = ttk.Frame(self.window, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.lbl_status = ttk.Label(
            self.main_frame,
            text="Initializing...",
            font=("Helvetica", 11)
        )
        self.lbl_status.pack(pady=(10, 10))

        self.progress_bar = ttk.Progressbar(
            self.main_frame,
            orient=tk.HORIZONTAL,
            length=300,
            mode='determinate'
        )
        self.progress_bar.pack(pady=(0, 20))

        # Disable the close button to prevent interruption during processing
        self.window.protocol("WM_DELETE_WINDOW", lambda: None)
        self.window.update()

    def center_window(self):
        """Calculates the screen center and positions the window accordingly."""
        self.window.update_idletasks()
        w = self.window.winfo_width()
        h = self.window.winfo_height()
        # Fallback dimensions if window is not yet rendered
        if w < 100: w = 400
        if h < 100: h = 150
        
        x = (self.window.winfo_screenwidth() // 2) - (w // 2)
        y = (self.window.winfo_screenheight() // 2) - (h // 2)
        self.window.geometry(f'{w}x{h}+{x}+{y}')

    def update(self, message, percent_complete):
        """
        Updates the GUI label/bar and prints the message to the terminal.

        Args:
            message (str): The status message to display.
            percent_complete (int): Integer between 0 and 100.
        """
        # 1. Update Terminal
        print(f"[Progress {percent_complete}%] {message}")

        # 2. Update GUI
        self.lbl_status.config(text=message)
        self.progress_bar['value'] = percent_complete
        self.window.update()

    def close(self):
        """Destroys the progress window."""
        self.window.destroy()


def generate_timestamp():
    """
    Generates a current Unix timestamp.
    
    Returns:
        int: The integer representation of the current time.
    """
    return int(time.time())


def escape_html(text):
    """
    Escapes HTML special characters to prevent broken markup.
    
    Args:
        text (str/obj): The input text to escape.
    
    Returns:
        str: Safe HTML string, or empty string if input is None/NaN.
    """
    if pd.isna(text) or text is None:
        return ""
    return html.escape(str(text))


def reveal_in_file_manager(file_path):
    """
    Reveals the specific file in the system's default file manager.
    Supports macOS (Finder), Windows (Explorer), and Linux (xdg-open).
    
    Args:
        file_path (str): The full path to the file.
    """
    try:
        if sys.platform == 'darwin':
            # macOS: 'open -R' highlights the file in Finder
            subprocess.run(['open', '-R', file_path], check=True)
        
        elif sys.platform == 'win32':
            # Windows: 'explorer /select,' highlights the file in Explorer
            # Normalize path (swap forward slashes for backslashes)
            norm_path = os.path.normpath(file_path)
            subprocess.run(['explorer', '/select,', norm_path], check=True)
            
        else:
            # Linux/Unix: 'xdg-open' usually opens the parent folder
            # Highlighting a specific file is desktop-env dependent (Nautilus vs Dolphin),
            # so opening the directory is the safest cross-distro fallback.
            parent_dir = os.path.dirname(file_path)
            subprocess.run(['xdg-open', parent_dir], check=True)
            
    except Exception as e:
        print(f"Error revealing file in manager: {e}")


def center_window(window):
    """
    Centers a Tkinter window on the primary screen.
    
    Args:
        window (tk.Tk or tk.Toplevel): The window object to center.
    """
    window.update_idletasks()
    width = window.winfo_reqwidth()
    height = window.winfo_reqheight()
    
    # Handle OS glitches where dimensions report as 1x1
    if width < 100: width = 500
    if height < 100: height = 300
    
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'+{x}+{y}')


def generate_html_recursive(data_structure, indent_level=1):
    """
    Recursively traverses the bookmark tree structure to generate Netscape HTML.
    """
    html_output = []
    indent = "    " * indent_level

    # Process Bookmarks in current node
    for bookmark in data_structure.get('_bookmarks_', []):
        ts = generate_timestamp()
        title = escape_html(bookmark.get('title'))
        url = escape_html(bookmark.get('url'))
        
        if not title and not url:
            continue
        if not title:
            title = "Untitled Bookmark"
        if not url:
            url = "#"
            
        html_output.append(
            f'{indent}<DT><A HREF="{url}" ADD_DATE="{ts}" LAST_MODIFIED="{ts}">{title}</A>'
        )

    # Process Sub-folders in current node
    for folder_name, sub_structure in data_structure.get('_folders_', {}).items():
        ts = generate_timestamp()
        escaped_folder_name = escape_html(folder_name)
        if not escaped_folder_name:
            escaped_folder_name = "Untitled Folder"
            
        html_output.append(
            f'{indent}<DT><H3 ADD_DATE="{ts}" LAST_MODIFIED="{ts}" '
            f'PERSONAL_TOOLBAR_FOLDER="false">{escaped_folder_name}</H3>'
        )
        html_output.append(f'{indent}<DL><p>')
        html_output.append(generate_html_recursive(sub_structure, indent_level + 1))
        
        # --- FIX IS ON THE LINE BELOW ---
        # WAS: html_output.append(f'{indent}<DL><p>') 
        # NOW: Changed to </DL> to properly close the list
        html_output.append(f'{indent}</DL><p>') 

    return "\n".join(html_output)


def build_bookmark_tree(df, title_col, url_col, dynamic_folder_cols):
    """
    Converts a flat Pandas DataFrame into a nested dictionary structure.

    Args:
        df (pd.DataFrame): The source DataFrame.
        title_col (str): Column name for titles.
        url_col (str): Column name for URLs.
        dynamic_folder_cols (list): List of column names representing folder depth.

    Returns:
        dict: A nested dictionary with keys '_bookmarks_' and '_folders_'.
    """
    tree = {'_bookmarks_': [], '_folders_': {}}

    for _, row in df.iterrows():
        current_level_node = tree
        path_folders_for_row = []

        # Extract folder path for this specific row
        for col_name in dynamic_folder_cols:
            folder_name_in_cell = row.get(col_name)
            if pd.notna(folder_name_in_cell) and str(folder_name_in_cell).strip() != "":
                path_folders_for_row.append(str(folder_name_in_cell).strip())
            else:
                # Stop at the first empty folder column
                break
        
        # Traverse/Create the tree structure based on path
        for folder in path_folders_for_row:
            if folder not in current_level_node['_folders_']:
                current_level_node['_folders_'][folder] = {
                    '_bookmarks_': [], 
                    '_folders_': {}
                }
            current_level_node = current_level_node['_folders_'][folder]
        
        # Add the bookmark to the final node
        bookmark_title_val = row.get(title_col)
        bookmark_url_val = row.get(url_col)
        has_title = pd.notna(bookmark_title_val) and str(bookmark_title_val).strip() != ""
        has_url = pd.notna(bookmark_url_val) and str(bookmark_url_val).strip() != ""
        
        if has_title or has_url:
            bm_data = {
                'title': str(bookmark_title_val).strip() if has_title else "Untitled Bookmark",
                'url': str(bookmark_url_val).strip() if has_url else "#"
            }
            current_level_node['_bookmarks_'].append(bm_data)
            
    return tree


def analyze_tree_stats(node, current_depth, stats_dict):
    """
    Traverses the generated tree to collect statistics for the summary report.

    Args:
        node (dict): The current tree node.
        current_depth (int): The current depth level.
        stats_dict (dict): Accumulator dictionary for stats.
    """
    stats_dict['total_bookmarks'] += len(node.get('_bookmarks_', []))
    for folder_name, sub_node in node.get('_folders_', {}).items():
        stats_dict['folders_per_level'][current_depth].add(folder_name)
        analyze_tree_stats(sub_node, current_depth + 1, stats_dict)


def get_summary_message(stats):
    """
    Formats the statistics dictionary into a readable string.

    Args:
        stats (dict): Dictionary containing 'total_bookmarks' and 'folders_per_level'.

    Returns:
        str: Formatted summary text.
    """
    lines = []
    lines.append(f"Total bookmarks processed: {stats['total_bookmarks']}")
    if not stats['folders_per_level']:
        lines.append("No folders were created.")
    else:
        lines.append("Folders created per level:")
        for depth, folders_set in sorted(stats['folders_per_level'].items()):
            level_name = f"L{depth + 1}"
            lines.append(f"  • {level_name}: {len(folders_set)} folder(s)")
    return "\n".join(lines)


# --- GUI HELPERS ---

def ask_header_row_configuration(parent_root):
    """
    Prompts the user to specify if the Excel headers are in Row 1 or Row 2.
    
    Returns:
        int: 0 for Row 1, 1 for Row 2.
    """
    config_win = tk.Toplevel(parent_root)
    config_win.title("Excel Setup")
    config_win.attributes('-topmost', True)
    config_win.lift()
    config_win.focus_force()
    
    selection = [-1]

    def select_row1():
        selection[0] = 0
        config_win.destroy()

    def select_row2():
        selection[0] = 1
        config_win.destroy()

    def on_close():
        config_win.destroy()
        sys.exit(0)

    config_win.protocol("WM_DELETE_WINDOW", on_close)

    main_frame = ttk.Frame(config_win, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(main_frame, text="Where are your Column Headers?", 
              font=("Helvetica", 14, "bold")).pack(pady=(0, 10))
    ttk.Label(main_frame, text="Does your file have headers in Row 1 or Row 2?").pack(pady=(0, 20))

    ttk.Button(main_frame, text="Row 1 (Standard)", command=select_row1).pack(pady=5, fill=tk.X)
    ttk.Button(main_frame, text="Row 2 (First row is blank/notes)", command=select_row2).pack(pady=5, fill=tk.X)

    center_window(config_win)
    parent_root.wait_window(config_win)
    parent_root.update()
    
    return selection[0]


def show_summary_window(summary_text_str, output_file_path, parent_root):
    """
    Displays the final summary window with stats and the output path.
    """
    try:
        summary_win = tk.Toplevel(parent_root)
        summary_win.title("Import Summary")
        summary_win.attributes('-topmost', True)
        summary_win.lift()
        
        main_frame = ttk.Frame(summary_win, padding="20")
        main_frame.pack(expand=True, fill=tk.BOTH)

        ttk.Label(main_frame, text="Conversion Complete!", 
                  font=("Helvetica", 16, "bold")).pack(pady=(0,15))
        
        ttk.Label(main_frame, text=f"Output file saved to:\n{output_file_path}", 
                  wraplength=550, justify=tk.LEFT).pack(pady=(0,15), anchor=tk.W, fill=tk.X)
        
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        summary_text_widget = tk.Text(main_frame, wrap=tk.WORD, height=8, 
                                      relief=tk.FLAT, borderwidth=0, highlightthickness=0)
        try: 
            summary_text_widget.config(background=main_frame.cget("background"))
        except: 
            pass
        
        summary_text_widget.insert(tk.END, summary_text_str)
        summary_text_widget.config(state=tk.DISABLED)
        summary_text_widget.pack(anchor=tk.W, fill=tk.BOTH, expand=True, pady=5)
        
        ttk.Button(main_frame, text="OK", command=summary_win.destroy).pack(pady=(15,0))
        
        center_window(summary_win)
        parent_root.wait_window(summary_win)
        
    except Exception as e:
        # Fallback if custom window fails
        messagebox.showinfo("Summary", summary_text_str, parent=parent_root)


def ask_for_excel_file(parent_root):
    """Opens a file dialog for Excel selection."""
    print("Selecting Input Excel file...")
    parent_root.update()
    file_path = filedialog.askopenfilename(
        parent=parent_root,
        title="Select Input Excel File (.xlsx, .xls)",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if not file_path:
        print("Selection Cancelled.")
        sys.exit(0)
    print(f"Input Selected: {file_path}")
    return file_path


def ask_for_output_html_file(default_filename, initial_dir, parent_root):
    """Opens a save dialog for the output HTML file."""
    print("Selecting Output location...")
    parent_root.update()
    output_path = filedialog.asksaveasfilename(
        parent=parent_root,
        title="Save Output Bookmarks HTML File As",
        defaultextension=".html",
        filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
        initialfile=default_filename,
        initialdir=initial_dir
    )
    if not output_path:
        print("Export cancelled.")
        sys.exit(0)
    return output_path


def main():
    """
    Main application entry point. Orchestrates the GUI setup, data processing,
    and file generation.
    """
    root = None
    try:
        # Initialize Hidden Root for Tkinter
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        root.lift()
        root.focus_force()
        root.update()
        
        # Apply Native Theme if available (macOS)
        style = ttk.Style()
        if 'aqua' in style.theme_names():
            style.theme_use('aqua')
        
        # 1. Configuration & File Selection
        rows_to_skip = ask_header_row_configuration(root)
        excel_file_path = ask_for_excel_file(root)
        
        # Prepare output paths
        input_basename = os.path.basename(excel_file_path)
        input_filestem, _ = os.path.splitext(input_basename)
        timestamp_str = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_output_filename = f"{input_filestem}_{timestamp_str}.html"
        initial_save_dir = os.path.dirname(excel_file_path)
        output_html_file = ask_for_output_html_file(default_output_filename, initial_save_dir, root)

        # 2. Start Progress Loader
        progress = ProgressLoader(root, title="NetscapeGen")
        
        # 3. Read Excel Data
        progress.update("Reading Excel data...", 10)
        try:
            time.sleep(0.5) # UI Refresh buffer
            df = pd.read_excel(excel_file_path, skiprows=rows_to_skip)
        except Exception as e:
            progress.close()
            messagebox.showerror("Excel Read Error", f"Error reading Excel file:\n{e}", parent=root)
            return

        # 4. Validate Columns
        progress.update("Validating columns...", 30)
        if TITLE_COLUMN not in df.columns or URL_COLUMN not in df.columns:
            progress.close()
            msg = f"Error: Required columns '{TITLE_COLUMN}' or '{URL_COLUMN}' not found."
            messagebox.showerror("Column Error", msg, parent=root)
            return

        # 5. Detect Dynamic Folder Structure
        progress.update("Analyzing structure...", 40)
        folder_column_candidates = [] 
        for col_name in df.columns.tolist():
            match = re.fullmatch(r'FolderL(\d+)', col_name) 
            if match:
                folder_column_candidates.append((int(match.group(1)), col_name))
        
        # Sort columns by level index (L1, L2, L3...)
        folder_column_candidates.sort(key=lambda x: x[0])
        dynamic_folder_columns = [name for num, name in folder_column_candidates]

        # 6. Clean Data
        progress.update("Cleaning data...", 50)
        not_nan_title = df[TITLE_COLUMN].notna() & (df[TITLE_COLUMN].astype(str).str.strip() != '')
        not_nan_url = df[URL_COLUMN].notna() & (df[URL_COLUMN].astype(str).str.strip() != '')
        df_cleaned = df[not_nan_title | not_nan_url].copy()

        # Sanitize folder columns (fill NaNs with empty strings)
        for col_name in dynamic_folder_columns:
            if col_name in df_cleaned.columns: 
                df_cleaned[col_name] = df_cleaned[col_name].fillna('').astype(str)

        # 7. Build Hierarchical Tree
        progress.update("Building bookmark tree...", 70)
        bookmark_tree = build_bookmark_tree(
            df_cleaned, 
            TITLE_COLUMN, 
            URL_COLUMN, 
            dynamic_folder_columns
        )
        
        # 8. Generate HTML Content
        progress.update("Generating HTML...", 85)
        html_header = (
            "<!DOCTYPE NETSCAPE-Bookmark-file-1>\n"
            "<META HTTP-EQUIV=\"Content-Type\" CONTENT=\"text/html; charset=UTF-8\">\n"
            "<TITLE>Bookmarks</TITLE>\n"
            "<H1>Bookmarks</H1>\n"
            "<DL><p>\n"
        )
        html_content = generate_html_recursive(bookmark_tree)
        html_footer = "</DL><p>"
        full_html = html_header + html_content + "\n" + html_footer

        # 9. Save and Reveal
        progress.update("Saving file...", 95)
        try:
            with open(output_html_file, 'w', encoding='utf-8') as f:
                f.write(full_html)
            reveal_in_file_manager(output_html_file)
        except Exception as e:
            progress.close()
            messagebox.showerror("File Write Error", f"Error writing file:\n{e}", parent=root)
            return

        # 10. Finalize
        import_stats = {
            'total_bookmarks': 0, 
            'folders_per_level': collections.defaultdict(set)
        }
        analyze_tree_stats(bookmark_tree, 0, import_stats) 
        summary_message_str = get_summary_message(import_stats)
        
        progress.update("Done!", 100)
        time.sleep(0.5) 
        progress.close()
        
        show_summary_window(summary_message_str, output_html_file, root)

    finally:
        if root:
            root.destroy()

if __name__ == "__main__":
    print("=" * 40)
    print("NetscapeGen")
    print("=" * 40)
    main()
    print("\n--- Script finished. ---")