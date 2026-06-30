# startup.py

import tkinter as tk
from tkinter import ttk
from pathlib import Path

try:
    from PIL import Image, ImageTk
    HAS_PILLOW = True
except ImportError:
    HAS_PILLOW = False

def show_startup_logo(root, image_candidates=("assets/logo.png", "assets/logo.gif", "assets/logo.ppm", "assets/image_cef6a4.jpg")):
    """
    Creates and displays a splash screen, using Pillow for robust image support.
    """
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.configure(bg="#f4f4f4")

    image_path = None
    for candidate in image_candidates:
        if Path(candidate).exists():
            image_path = candidate
            break

    splash.image_ref = None

    if image_path:
        try:
            if HAS_PILLOW:
                img_open = Image.open(image_path)
                
                # --- NEW: Shrink the image if it's too big! ---
                max_size = 350
                if img_open.width > max_size or img_open.height > max_size:
                    # thumbnail scales it down while keeping the correct aspect ratio
                    img_open.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
                # ----------------------------------------------
                
                img = ImageTk.PhotoImage(img_open)
            else:
                img = tk.PhotoImage(file=image_path)
            
            splash.image_ref = img 
            lbl = tk.Label(splash, image=img, borderwidth=0, highlightthickness=0, bg="#f4f4f4")
            lbl.pack(padx=20, pady=(20, 0))
        except Exception as e:
            lbl = tk.Label(splash, text="Loading…", font=("Helvetica", 14), bg="#f4f4f4")
            lbl.pack(padx=20, pady=20)
    else:
        lbl = tk.Label(splash, text="Loading…", font=("Helvetica", 14), bg="#f4f4f4")
        lbl.pack(padx=20, pady=20)

    # --- Progress Bar and Status Label ---
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(splash, variable=progress_var, maximum=100, length=250)
    progress_bar.pack(fill='x', padx=40, pady=(15, 5))
    
    status_label = tk.Label(splash, text="Waking up...", font=("Helvetica", 9), fg="#555555", bg="#f4f4f4")
    status_label.pack(pady=(0, 15))

    # Let the UI engine calculate the layout with the resized image
    splash.update_idletasks()
    
    w = splash.winfo_reqwidth()
    h = splash.winfo_reqheight()
    
    sw = splash.winfo_screenwidth()
    sh = splash.winfo_screenheight()
    
    x = (sw - w) // 2
    y = (sh - h) // 2
    
    splash.geometry(f"{w}x{h}+{x}+{y}")
    
    return splash, progress_var, status_label