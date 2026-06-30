# cache_generator.py
import os
import logging
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- Robust Logging Configuration ---
# Create a custom logger specifically for this module
logger = logging.getLogger("cache_generator")
logger.setLevel(logging.INFO) 

# Only add handlers if they haven't been added yet (prevents duplicate log lines)
if not logger.handlers:
    # 1. Format the logs to include the thread name
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(threadName)s: %(message)s")
    
    # 2. Setup Console output
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # 3. Setup File output
    file_handler = logging.FileHandler("cache_generator.log", mode="w", encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # Prevent log messages from bubbling up to the root logger
    logger.propagate = False

# Constants for caching logic
IMAGE_EXTENSIONS = [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tiff"]
SUPPORTED_EXTENSIONS = [".txt", ".pdf", ".docx", ".xls", ".html"] + IMAGE_EXTENSIONS

def extract_text(path):
    """
    Extracts text content from a supported file type.
    Uses LAZY IMPORTS to prevent startup lag across threads.
    """
    ext = os.path.splitext(path)[1].lower()
    logger.debug(f"Starting extraction for: {path} (Type: {ext})")
    
    try:
        if ext == ".txt":
            return Path(path).read_text(errors="ignore")
            
        elif ext == ".pdf":
            import fitz  # PyMuPDF Lazy Import
            with fitz.open(path) as doc:
                return " ".join(page.get_text() for page in doc)
                
        elif ext == ".docx":
            import docx  # python-docx Lazy Import
            doc = docx.Document(path)
            return " ".join(p.text for p in doc.paragraphs)
            
        elif ext == ".xls":
            import xlrd  # xlrd Lazy Import
            book = xlrd.open_workbook(path)
            return " ".join(
                str(cell.value)
                for sheet in book.sheets()
                for row in range(sheet.nrows)
                for cell in sheet.row(row)
            )
            
        elif ext == ".html":
            from bs4 import BeautifulSoup  # BeautifulSoup Lazy Import
            return BeautifulSoup(Path(path).read_text(errors="ignore"), "html.parser").get_text()
            
        elif ext in IMAGE_EXTENSIONS:
            from PIL import Image  # Pillow Lazy Import
            # --- Image Metadata Extraction Logic ---
            extracted_text = []
            with Image.open(path) as img:
                if 'Description' in img.info:
                    extracted_text.append(str(img.info['Description']))
                if 'Title' in img.info:
                    extracted_text.append(str(img.info['Title']))

                try:
                    exif = img.getexif()
                    if exif:
                        if 270 in exif:
                            val = exif[270]
                            if isinstance(val, bytes):
                                val = val.decode('utf-8', errors='ignore')
                            extracted_text.append(str(val))
                        
                        if 40091 in exif:
                            val = exif[40091]
                            if isinstance(val, bytes):
                                val = val.decode('utf-16le', errors='ignore').rstrip('\x00')
                            extracted_text.append(str(val))
                except Exception as e:
                    logger.debug(f"No EXIF data found or readable in {path}: {e}")
                    pass 
                    
            unique_text = list(dict.fromkeys(extracted_text))
            final_text = "\n\n".join(unique_text).strip()
            
            return final_text

    except Exception as e:
        logger.error(f"Failed to extract text from {path}: {e}", exc_info=False)
    return ""

def _process_single_file(path, old_index):
    """
    Helper function to process a single file. Checks mtime and extracts text if needed.
    """
    try:
        mtime = os.path.getmtime(path)
    except OSError as e:
        logger.warning(f"Could not retrieve mtime for {path}: {e}")
        mtime = None

    if path in old_index and old_index[path][1] == mtime:
        logger.debug(f"Cache hit: Skipping extraction for {path}")
        text = old_index[path][0]
    else:
        logger.debug(f"Cache miss/update: Extracting text for {path}")
        text = extract_text(path)

    return (path, mtime, text)

def build_cache(directory, cache, force_rebuild=False, progress_callback=None, max_workers=None):
    """
    Builds or updates the cache for a single directory using parallel Thread processing.
    """
    logger.info(f"Starting cache build for directory: {directory} (Force Rebuild: {force_rebuild})")
    
    old = cache.get(directory, {})
    old_files = old.get("files", [])
    old_texts = old.get("texts", [])
    old_mtimes = old.get("mtimes", [])
    
    if len(old_mtimes) != len(old_files):
        logger.warning("Cache mismatch detected (mtimes length != files length). Forcing rebuild.")
        force_rebuild = True

    old_index = {} if force_rebuild else {
        str(Path(fp).resolve()): (txt, mtime)
        for fp, txt, mtime in zip(old_files, old_texts, old_mtimes)
    }

    filepaths = [
        str(Path(p).resolve()) for p in Path(directory).rglob("*")
        if p.suffix.lower() in SUPPORTED_EXTENSIONS
    ]
    total = len(filepaths)
    logger.info(f"Found {total} supported files to process.")
    
    new_entries = []

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_process_single_file, path, old_index): path for path in filepaths}
        
        for idx, future in enumerate(as_completed(futures), start=1):
            try:
                result = future.result()
                new_entries.append(result)
            except Exception as e:
                path = futures[future]
                logger.error(f"Task for {path} failed during parallel execution: {e}", exc_info=True)
            
            if progress_callback:
                pct = int(idx / total * 100) if total else 100
                progress_callback(f"Processing {idx}/{total} ({pct}%)")
                if pct % 10 == 0:  
                    logger.info(f"Progress: {idx}/{total} ({pct}%) completed.")
    
    new_entries.sort(key=lambda x: x[0])

    if new_entries:
        files, mtimes, texts = zip(*new_entries)
    else:
        files, mtimes, texts = [], [], []

    logger.info(f"Cache build complete. Processed {len(files)} files successfully.")

    return {
        "files": list(files),
        "texts": list(texts),
        "mtimes": list(mtimes),
        "timestamp": datetime.now().isoformat()
    }