import os
import re
import time
import logging
import joblib
import scipy.sparse as sp
from pathlib import Path
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# --- Logging Configuration ---
logging.basicConfig(
    filename='recommender_profiling.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filemode='w' 
)
logger = logging.getLogger(__name__)

CACHE_DIR = Path(__file__).parent.parent / "cache"
CACHE_DIR.mkdir(exist_ok=True)
MODEL_PATH = CACHE_DIR / "recommender_model.pkl"

# --- Global State for TF-IDF Caching ---
_model_cache = {
    "vectorizer": None,
    "corpus_matrix": None,
    "vocab": None,
    "file_map": [],
    "candidate_texts": [],
    "directory_map": [],
    "mtime_map": {} 
}

def load_model():
    """Loads the model from disk using joblib for extreme speed."""
    global _model_cache
    if _model_cache.get("vectorizer") is not None:
        return True
    if MODEL_PATH.exists():
        try:
            # joblib reads SciPy sparse matrices instantly
            disk_cache = joblib.load(MODEL_PATH)
            _model_cache.update(disk_cache)
            logger.info("Successfully loaded Recommender Model from disk.")
            return True
        except Exception as e:
            logger.error(f"Failed to load Recommender Model: {e}")
    return False

def update_model(cache):
    """
    Incrementally builds or updates the TF-IDF matrix.
    Called strictly when 'Reload All' is executed in the UI.
    """
    global _model_cache
    logger.info("=== Starting Recommender Model Update ===")
    total_start = time.time()

    has_old_model = load_model()
    
    if not has_old_model:
        logger.info("No existing model found. Doing FULL TF-IDF Vectorization...")
        
        target_texts, target_files, target_dirs, target_mtimes = [], [], [], []
        for directory, data in cache.items():
            texts = data.get("texts", [])
            target_texts.extend(texts)
            target_files.extend(data.get("files", []))
            target_mtimes.extend(data.get("mtimes", []))
            target_dirs.extend([directory] * len(texts))

        if not target_texts:
            logger.warning("No texts found in cache to train on.")
            return

        vect = TfidfVectorizer()
        corpus_matrix = vect.fit_transform(target_texts)
        
        _model_cache["vectorizer"] = vect
        _model_cache["corpus_matrix"] = corpus_matrix
        _model_cache["vocab"] = vect.vocabulary_
        _model_cache["file_map"] = target_files
        _model_cache["candidate_texts"] = target_texts
        _model_cache["directory_map"] = target_dirs
        _model_cache["mtime_map"] = dict(zip(target_files, target_mtimes))
    else:
        logger.info("Existing model found. Doing FAST INCREMENTAL TF-IDF update...")
        old_vect = _model_cache["vectorizer"]
        old_matrix = _model_cache["corpus_matrix"]
        old_file_to_idx = {f: i for i, f in enumerate(_model_cache["file_map"])}
        old_mtimes_map = _model_cache.get("mtime_map", {})
        
        keep_indices = []
        final_files, final_texts, final_dirs, final_mtimes = [], [], [], {}
        new_files, new_texts, new_dirs, new_mtimes = [], [], [], []

        # 1. Categorize files into "Keep" vs "New/Modified"
        for directory, data in cache.items():
            for f, txt, mtime in zip(data.get("files", []), data.get("texts", []), data.get("mtimes", [])):
                if f in old_file_to_idx and old_mtimes_map.get(f) == mtime:
                    keep_indices.append(old_file_to_idx[f])
                    final_files.append(f)
                    final_texts.append(txt)
                    final_dirs.append(directory)
                    final_mtimes[f] = mtime
                else:
                    new_texts.append(txt)
                    new_files.append(f)
                    new_dirs.append(directory)
                    new_mtimes.append(mtime)
        
        logger.info(f"Reusing {len(keep_indices)} unchanged documents. Vectorizing {len(new_texts)} new/changed documents.")
        
        # 2. FAST Sparse Matrix slicing for kept files (One operation!)
        kept_matrix = old_matrix[keep_indices, :] if keep_indices else None
        
        # 3. Transform only the new texts
        if new_texts:
            step_start = time.time()
            new_matrix = old_vect.transform(new_texts)
            logger.info(f"Transforming {len(new_texts)} new texts took {time.time() - step_start:.4f}s")
            
            final_files.extend(new_files)
            final_texts.extend(new_texts)
            final_dirs.extend(new_dirs)
            for f, m in zip(new_files, new_mtimes):
                final_mtimes[f] = m
        else:
            new_matrix = None
            
        # 4. FAST Matrix Merging (One operation!)
        if kept_matrix is not None and new_matrix is not None:
            new_corpus_matrix = sp.vstack([kept_matrix, new_matrix])
        elif kept_matrix is not None:
            new_corpus_matrix = kept_matrix
        elif new_matrix is not None:
            new_corpus_matrix = new_matrix
        else:
            new_corpus_matrix = sp.csr_matrix((0,0))
            
        _model_cache["corpus_matrix"] = new_corpus_matrix
        _model_cache["file_map"] = final_files
        _model_cache["candidate_texts"] = final_texts
        _model_cache["directory_map"] = final_dirs
        _model_cache["mtime_map"] = final_mtimes

    # --- SAVE INCREMENTAL STATE TO DISK ---
    try:
        # Create a shallow copy and drop the raw texts before saving
        cache_to_save = _model_cache.copy()
        cache_to_save["candidate_texts"] = [] 
        
        joblib.dump(cache_to_save, MODEL_PATH)
        logger.info(f"Model saved to {MODEL_PATH} using joblib")
    except Exception as e:
        logger.error(f"Failed to save model to disk: {e}")

    logger.info(f"=== Recommender Model Update Complete in {time.time() - total_start:.4f}s ===")


def get_recommendations(current_text, cache, directory_active_status, amplifiers, silencers, rest_part, word_count_threshold, length_factor):
    global _model_cache
    total_start = time.time()
    logger.info("=== Starting get_recommendations ===")

    # Do NOT rebuild the matrix automatically. Just load from disk if missing.
    if not load_model():
        logger.warning("Recommender model is not trained. Please click 'Reload All'.")
        return []

    # --- 3. Query Vectorization & Basic Amplification ---
    step_start = time.time()
    
    try:
        query_matrix = _model_cache["vectorizer"].transform([current_text])
        query_vec = query_matrix.toarray()
    except Exception as e:
        logger.error(f"Failed to transform query text: {e}")
        return []
    
    F = amplifiers.get("factor", 2.0)
    vocab = _model_cache["vocab"]
    for w in amplifiers.get("words", []):
        key = w.lower()
        if key in vocab:
            col = vocab[key]
            query_vec[0, col] *= F
            
    logger.info(f"[Step 3] Query Transform & Amplification: {time.time() - step_start:.4f}s")

    # --- 4. Cosine Similarity ---
    step_start = time.time()
    similarities = cosine_similarity(query_vec, _model_cache["corpus_matrix"]).ravel()
    logger.info(f"[Step 4] Cosine Similarity calculation: {time.time() - step_start:.4f}s")

    # Rebuild candidate texts from live cache if missing from memory cache
    if not _model_cache.get("candidate_texts"):
        logger.info("Rebuilding text mappings from live cache...")
        file_to_text = {}
        for directory, data in cache.items():
            for f, txt in zip(data.get("files", []), data.get("texts", [])):
                file_to_text[f] = txt
        _model_cache["candidate_texts"] = [file_to_text.get(f, "") for f in _model_cache["file_map"]]

    candidate_texts = _model_cache["candidate_texts"]
    file_map = _model_cache["file_map"]
    directory_map = _model_cache["directory_map"]

    # --- 5. Base Ranking & Candidate Slicing ---
    step_start = time.time()
    base_scores = []
    
    # Fast loop: filter out inactive directories and zero-score documents
    for sim, text, path, directory in zip(similarities, candidate_texts, file_map, directory_map):
        if directory_active_status.get(directory, True) and float(sim) > 0.0:
            base_scores.append((float(sim), text, path))
            
    # SLICE THE POOL: Take only the top 150 documents for expensive regex evaluation
    POOL_SIZE = 150
    candidate_pool = sorted(base_scores, key=lambda x: x[0], reverse=True)[:POOL_SIZE]
    logger.info(f"[Step 5] Sliced Top {len(candidate_pool)} documents for Regex processing: {time.time() - step_start:.4f}s")

    # --- 6. Pre-computation of Regexes ---
    amp_words = [w.lower() for w in amplifiers.get("words", []) if w]
    amp_pattern = re.compile(rf"\b({'|'.join(map(re.escape, amp_words))})\b") if amp_words else None

    silencer_words = [s.strip().lower() for s in silencers.get("words", []) if s.strip()]
    silencer_text_pattern = re.compile(rf"\b({'|'.join(map(re.escape, silencer_words))})\b") if silencer_words else None

    amp_factor = amplifiers.get("factor", 2.0)
    sil_factor = float(silencers.get("factor", 1.0))
    rest_lower = (rest_part or "").lower()

    # --- 7. Target Adjustments Loop (Only on Top 150) ---
    step_start = time.time()
    
    # PASS 1: Calculate Keyword Diversity scores for the candidate pool to find the maximum limit
    doc_kw_scores = []
    max_kw_score = 0.0

    if amp_pattern:
        for sim, text, path in candidate_pool:
            lower_text = (text or "").lower()
            lower_path = path.lower()
            
            all_matches = amp_pattern.findall(lower_text) + amp_pattern.findall(lower_path)
            
            if all_matches:
                T = len(all_matches)
                U = len(set(all_matches))
                score = U + (T / (T + 1))
                doc_kw_scores.append(score)
                if score > max_kw_score:
                    max_kw_score = score
            else:
                doc_kw_scores.append(0.0)
    else:
        doc_kw_scores = [0.0] * len(candidate_pool)

    # PASS 2: Apply Final Multipliers & Length Penalties
    adjusted_scores = []
    for i, (base_sim, text, path) in enumerate(candidate_pool):
        adj = base_sim
        doc_multiplier = 1.0
        lower_path = path.lower()
        lower_text = (text or "").lower()

        # 1. Apply Strict rest_part Matching if used
        if rest_lower:
            if rest_lower in lower_path:
                doc_multiplier *= (amp_factor * 300)
            if rest_lower in lower_text:
                rest_hits = lower_text.count(rest_lower)
                if rest_hits > 0:
                    doc_multiplier *= max(0.001, 1 + ((amp_factor * 300) - 1) * rest_hits)
        
        # 2. Apply Proportional Keyword Diversity Amplification
        if max_kw_score > 0.0 and doc_kw_scores[i] > 0.0:
            # Scaled linearly so the absolute best document hits 'amp_factor' exactly
            scaled_amp = 1.0 + (amp_factor - 1.0) * (doc_kw_scores[i] / max_kw_score)
            doc_multiplier *= scaled_amp

        # 3. Check Silencer Matches
        if silencer_words and silencer_text_pattern:
            silencer_matches = set(silencer_text_pattern.findall(lower_text))
            sil_hits = len(silencer_matches)
            
            for s in silencer_words:
                if s not in silencer_matches and s in lower_path:
                    sil_hits += 1
            
            if sil_hits:
                doc_multiplier *= (sil_factor ** sil_hits)

        adj *= doc_multiplier

        # 4. FAST WORD COUNT Length Penalty
        if ((text or "").count(" ") + 1) > word_count_threshold:
            adj *= length_factor
            
        adjusted_scores.append((adj, path))
        
    logger.info(f"[Step 7] Advanced Adjustments Loop on {len(candidate_pool)} documents: {time.time() - step_start:.4f}s")

    # --- 8. Final Ranking and Return ---
    step_start = time.time()
    ranked = sorted(adjusted_scores, key=lambda x: x[0], reverse=True)[:12]
    logger.info(f"[Step 8] Final Sorting & Selection: {time.time() - step_start:.4f}s")
    
    logger.info(f"=== Total get_recommendations execution: {time.time() - total_start:.4f}s ===")
    return ranked