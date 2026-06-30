# ocr_extractor.py
import os
import logging

try:
    import easyocr
except ImportError:
    easyocr = None

try:
    from PIL import Image, PngImagePlugin, ImageOps
    import piexif
    import numpy as np
except ImportError:
    Image = None
    piexif = None
    ImageOps = None
    np = None

logger = logging.getLogger(__name__)

# Cache the reader so we don't reload the AI model every time an image is clicked
_reader = None

def get_reader():
    global _reader
    if _reader is None and easyocr is not None:
        logger.info("Initializing EasyOCR Model on CPU...")
        # Explicitly set gpu=False to ensure a lean, CPU-only footprint
        _reader = easyocr.Reader(['en'], gpu=False)
    return _reader

def smart_join_text(results):
    """Joins lines smartly based on punctuation."""
    if not results:
        return ""

    final_text = ""
    sentence_enders = ('.', '!', '?', ':', ';', '"', "'")

    for i, line in enumerate(results):
        line = line.strip()
        if not line:
            continue
            
        final_text += line
        
        if i < len(results) - 1:
            if line.endswith(sentence_enders):
                final_text += "\n\n"
            else:
                final_text += " "
                
    return final_text

def embed_metadata(image_path, text):
    """Embeds text into the image metadata (EXIF for JPEG, tEXt chunk for PNG)."""
    try:
        ext = os.path.splitext(image_path)[1].lower()
        
        if ext == '.png':
            with Image.open(image_path) as img:
                metadata = PngImagePlugin.PngInfo()
                metadata.add_text("Description", text)
                img.save(image_path, "PNG", pnginfo=metadata)
                
        elif ext in ['.jpg', '.jpeg']:
            try:
                exif_dict = piexif.load(image_path)
            except Exception:
                exif_dict = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}, "thumbnail": None}
            
            exif_dict["0th"][piexif.ImageIFD.ImageDescription] = text.encode('utf-8')
            
            try:
                # Attempt to dump the modified original EXIF data
                exif_bytes = piexif.dump(exif_dict)
            except ValueError as e:
                logger.warning(f"piexif failed to dump original EXIF (likely due to MakerNotes). Generating clean EXIF. Error: {e}")
                
                # Fallback: Smartphone MakerNotes often break piexif serialization. 
                # Create a completely clean EXIF dictionary to bypass the junk data.
                clean_exif = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}, "thumbnail": None}
                
                # Preserve the orientation so the image doesn't rotate sideways!
                if "0th" in exif_dict and piexif.ImageIFD.Orientation in exif_dict["0th"]:
                    clean_exif["0th"][piexif.ImageIFD.Orientation] = exif_dict["0th"][piexif.ImageIFD.Orientation]
                    
                # Add our extracted text
                clean_exif["0th"][piexif.ImageIFD.ImageDescription] = text.encode('utf-8')
                
                # Dump the clean version
                exif_bytes = piexif.dump(clean_exif)
                
            piexif.insert(exif_bytes, image_path)
            
        else:
            logger.warning(f"Metadata embedding not supported for format {ext}")
            return False
            
        return True
    except Exception as e:
        logger.warning(f"Could not embed metadata in {os.path.basename(image_path)}: {e}")
        return False

def process_single_image(image_path):
    """
    Extracts text (both upright and rotated) and embeds it into the image. 
    Returns a tuple: (success_boolean, extracted_text_or_error_message)
    """
    if easyocr is None or Image is None or piexif is None:
        return False, "Missing required libraries. Run: pip install easyocr Pillow piexif numpy"

    try:
        reader = get_reader()
        
        with Image.open(image_path) as pil_img:
            # 1. Correct EXIF orientation first
            pil_img = ImageOps.exif_transpose(pil_img)
            
            # OPTIMIZATION 1: Downscale massive photos. 
            # 4000+ pixel images choke OCR engines. 2048px is the sweet spot 
            # where you retain enough detail for text, but process 4x faster.
            max_dim = 2048
            if max(pil_img.size) > max_dim:
                pil_img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
                
            # OPTIMIZATION 2: Convert to Grayscale early.
            # EasyOCR does this internally anyway, but doing it here reduces the 
            # numpy array memory size by 3x before handing it to the GPU/CPU.
            gray_img = pil_img.convert('L')
            
            # Prepare our two views
            img_array_upright = np.array(gray_img)
            img_array_rotated = np.array(gray_img.rotate(90, expand=True))
            
        # OPTIMIZATION 3: Relaxed OCR parameters.
        # We lowered canvas_size and mag_ratio because the image is already
        # properly scaled, preventing the AI from doing unnecessary upsampling.
        ocr_kwargs = {
            'detail': 0,
            'canvas_size': 2560, 
            'mag_ratio': 1.2,    
            'contrast_ths': 0.1,  
            'text_threshold': 0.5 
        }

        # --- THE DOUBLE RUN ---
        results_upright = reader.readtext(img_array_upright, **ocr_kwargs)
        results_rotated = reader.readtext(img_array_rotated, **ocr_kwargs)

        # Process text
        text_upright = smart_join_text(results_upright)
        text_rotated = smart_join_text(results_rotated)
        
        # Combine the results cleanly
        extracted_parts = []
        if text_upright.strip():
            extracted_parts.append("--- Upright Text ---\n" + text_upright)
        if text_rotated.strip():
            extracted_parts.append("--- Sideways/Spine Text ---\n" + text_rotated)
            
        extracted_text = "\n\n".join(extracted_parts).strip()
        
        if not extracted_text:
            return False, "No text detected in image."
            
        # Overwrite metadata
        if embed_metadata(image_path, extracted_text):
            return True, extracted_text
        else:
            return False, "Text extracted, but failed to embed metadata."
            
    except Exception as e:
        logger.error(f"ERROR processing {image_path}: {e}")
        return False, str(e)