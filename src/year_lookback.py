# --- year_lookback.py ---
import os
import datetime
import re
import calendar
import collections
import math
import fitz  # PyMuPDF
from pathlib import Path

# --- AI Import ---
try:
    from transformers import pipeline
    HAS_TRANSFORMERS = True
except ImportError:
    HAS_TRANSFORMERS = False
    print("[WARNING] 'transformers' library not found. Falling back to templates.")

# --- Configuration ---
STOP_WORDS = {
    # English
    "the", "be", "to", "of", "and", "a", "in", "that", "have", "i", "it", "for",
    "not", "on", "with", "he", "as", "you", "do", "at", "this", "but", "his",
    "by", "from", "they", "we", "say", "her", "she", "or", "an", "will", "my",
    "one", "all", "would", "there", "their", "what", "so", "up", "out", "if",
    "about", "who", "get", "which", "go", "me", "when", "make", "can", "like",
    "time", "no", "just", "him", "know", "take", "people", "into", "year", "your",
    
    # Portuguese
    "de", "a", "o", "que", "e", "do", "da", "em", "um", "para", "é", "com", "não", "uma", "os", "no", 
    "se", "na", "por", "mais", "as", "dos", "como", "mas", "foi", "ao", "ele", "das", "tem", "à", 
    
    # Generic / Numbers
    "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "http", "https", "www", "com", "org", "pdf", "txt"
}

THEMES = [
    {"bg": (0.1, 0.1, 0.1), "text": (0.9, 0.9, 0.9), "accent": (1, 0.4, 0.8), "sec": (0.2, 0.6, 1.0)}, 
]

# --- Global Generator Cache ---
_AI_GENERATOR = None

def _get_ai_generator():
    global _AI_GENERATOR
    if _AI_GENERATOR is None and HAS_TRANSFORMERS:
        try:
            _AI_GENERATOR = pipeline("text2text-generation", model="google/flan-t5-base")
        except Exception:
            _AI_GENERATOR = None
    return _AI_GENERATOR

def parse_years(years_input):
    if not years_input: return [datetime.datetime.now().year]
    years_input = str(years_input).strip()
    
    if '-' in years_input:
        try:
            parts = years_input.split('-')
            start, end = int(parts[0]), int(parts[1])
            step = 1 if start <= end else -1
            return list(range(start, end + step, step))
        except ValueError: pass 
            
    if ',' in years_input:
        try: return [int(y.strip()) for y in years_input.split(',')]
        except ValueError: pass
            
    try: return [int(years_input)]
    except ValueError: return [datetime.datetime.now().year]

def generate_report(cache, directory_active_status, save_dir, years_input=None):
    if years_input is None:
        try:
            import tkinter as tk
            from tkinter import simpledialog
            root = tk.Tk()
            root.withdraw() 
            root.attributes('-topmost', True) 
            user_in = simpledialog.askstring("Year Selection", "Enter years to generate (e.g., '2025', '2025-2022'):")
            root.destroy() 
            years_input = user_in if user_in else str(datetime.datetime.now().year)
        except Exception:
            years_input = str(datetime.datetime.now().year)

    years_to_process = parse_years(years_input)
    generated_files = []
    
    if HAS_TRANSFORMERS: _get_ai_generator()

    for year in years_to_process:
        pdf_path = _generate_single_year_report(cache, directory_active_status, save_dir, year)
        if pdf_path: generated_files.append(pdf_path)
    
    return generated_files[0] if generated_files else None

def _generate_single_year_report(cache, directory_active_status, save_dir, target_year):
    data_bundle = _gather_statistics(cache, directory_active_status, target_year)
    stats = data_bundle["stats"]
    
    if stats["total_docs"] == 0:
        return None

    pdf_filename = f"{target_year} Lookback.pdf"
    pdf_path = Path(save_dir) / pdf_filename
    doc = fitz.open()

    _draw_slide_title(doc, target_year, stats)
    _draw_slide_volume(doc, stats)
    _draw_slide_prime_time(doc, stats)
    _draw_slide_landscape(doc, stats, "Your Landscape", "Where you spent your time")
    _draw_slide_aura_weekly(doc, data_bundle["weekly_data"], "Your Aura", target_year)
    _draw_slide_heatmap(doc, target_year, stats)

    try:
        doc.save(str(pdf_path))
        return str(pdf_path)
    except Exception:
        return None

def _init_stats():
    return {
        "total_docs": 0,
        "total_words": 0,
        "longest_file": ("None", 0),
        "months": collections.defaultdict(int),
        "weekdays": collections.defaultdict(int),
        "hours": collections.defaultdict(int),
        "daily_counts": collections.defaultdict(int),
        "deep_folder_activity": {}, 
    }

def _gather_statistics(cache, active_status, year):
    stats = _init_stats()
    weekly_data = collections.defaultdict(lambda: {"words": collections.Counter(), "files": {}})
    date_pattern = re.compile(r'^(\d{4})(\d{2})(\d{2})')

    for directory, data in cache.items():
        if not active_status.get(directory, True): continue

        files = data.get("files", [])
        texts = data.get("texts", [])
        mtimes = data.get("mtimes", [])
        folder_name = os.path.basename(os.path.normpath(directory))

        for fpath, ftext, fmtime in zip(files, texts, mtimes):
            fname = os.path.basename(fpath)
            dt_obj = None
            
            match = date_pattern.match(fname)
            if match and int(match.group(1)) == year:
                dt_obj = datetime.datetime(year, int(match.group(2)), int(match.group(3)))
            
            if not dt_obj:
                try:
                    temp_dt = datetime.datetime.fromtimestamp(fmtime)
                    if temp_dt.year == year: dt_obj = temp_dt
                except Exception: pass

            if dt_obj:
                text_lower = ftext.lower()
                words = re.findall(r'\w+', text_lower)
                wc = len(words)

                stats["total_docs"] += 1
                stats["total_words"] += wc
                if wc > stats["longest_file"][1]:
                    stats["longest_file"] = (fname, wc)

                stats["months"][dt_obj.month] += 1
                stats["weekdays"][dt_obj.weekday()] += 1
                stats["hours"][dt_obj.hour] += 1
                stats["daily_counts"][(dt_obj.month, dt_obj.day)] += 1

                try:
                    rel_path = os.path.relpath(fpath, directory)
                    parts = rel_path.split(os.sep)
                    proj_name = f"{parts[0]}/{parts[1]}" if len(parts) > 2 else (parts[0] if len(parts) == 2 else folder_name)
                    
                    if proj_name not in stats["deep_folder_activity"]:
                        stats["deep_folder_activity"][proj_name] = {"count": 0, "months": collections.defaultdict(int)}
                    
                    stats["deep_folder_activity"][proj_name]["count"] += 1
                    stats["deep_folder_activity"][proj_name]["months"][dt_obj.month] += 1
                except ValueError: pass

                wk = dt_obj.isocalendar()[1]
                sig_words = [w for w in words if w not in STOP_WORDS and len(w) > 3 and not w.isdigit()]
                weekly_data[wk]["words"].update(sig_words)
                weekly_data[wk]["files"][fname] = sig_words

    return { "stats": stats, "weekly_data": weekly_data }

# --- Visualization Helpers ---
def _get_page(doc, theme_idx=0):
    page = doc.new_page(width=600, height=800)
    theme = THEMES[theme_idx % len(THEMES)]
    page.draw_rect(page.rect, color=theme["bg"], fill=theme["bg"])
    return page, theme

def draw_text_safe(page, x, y, text, fontsize, color, fontname="helv", align=None):
    try:
        if align == 1: 
            try:
                width = fitz.Font(fontname).text_length(text, fontsize)
                x = (page.rect.width - width) / 2
            except:
                x = (page.rect.width - (len(text) * fontsize * 0.5)) / 2
        page.insert_text((x, y + fontsize), str(text), fontsize=fontsize, fontname=fontname, color=color)
    except Exception:
        page.insert_text((x, y + fontsize), str(text), fontsize=fontsize, color=color)

def _draw_slide_title(doc, year, stats):
    page, theme = _get_page(doc, 0)
    page.draw_circle((300, 400), 150, color=theme["accent"], width=5)
    draw_text_safe(page, 0, 80, f"{year}", 80, theme["text"], "helvbo", align=1)
    draw_text_safe(page, 0, 50, "WRAPPED", 50, theme["text"], "helv", align=1)
    draw_text_safe(page, 0, 600, f"You worked on {stats['total_docs']} files this year.", 18, theme["text"], "helv", align=1)

def _draw_slide_volume(doc, stats):
    page, theme = _get_page(doc, 0)
    draw_text_safe(page, 0, 100, "The Output", 30, theme["accent"], "helvbo", align=1)
    
    draw_text_safe(page, 0, 150, f"{stats['total_words']:,}", 70, theme["text"], "helvbo", align=1)
    draw_text_safe(page, 0, 240, "Total Words Touched", 16, theme["text"], "helv", align=1)
    
    fname, wc = stats['longest_file']
    draw_text_safe(page, 0, 350, "Biggest File", 18, theme["sec"], "helvbo", align=1)
    
    disp_name = fname if len(fname) < 40 else fname[:37] + "..."
    draw_text_safe(page, 0, 380, f"{disp_name} ({wc:,} w)", 14, theme["text"], "helv", align=1)

def _draw_slide_prime_time(doc, stats):
    page, theme = _get_page(doc, 0)
    draw_text_safe(page, 0, 100, "Your Prime Time", 30, theme["text"], "helvbo", align=1)

    def get_peak_time(stats_hours):
        if not stats_hours: return "N/A"
        valid_hours = {h: c for h, c in stats_hours.items() if h != 0}
        if not valid_hours: return "Midnight"
        times = {
            "Morning": sum(valid_hours.get(h,0) for h in range(5, 12)),
            "Afternoon": sum(valid_hours.get(h,0) for h in range(12, 18)),
            "Evening": sum(valid_hours.get(h,0) for h in range(18, 24)),
            "Late Night": sum(valid_hours.get(h,0) for h in range(1, 5))
        }
        return max(times, key=times.get)

    draw_text_safe(page, 0, 200, "When you are most active", 18, theme["accent"], "helvbo", align=1)
    draw_text_safe(page, 0, 240, f"Peak: {get_peak_time(stats['hours'])}", 16, theme["text"], "helv", align=1)
    
    valid = [c for h, c in stats["hours"].items() if h != 0]
    max_h = max(valid) if valid else 1
    y_base = 350

    for h in range(24):
        if h == 0: continue 
        val = stats["hours"][h]
        bar_h = (val/max_h) * 150 if max_h > 0 else 0
        x = 50 + (h * 20)
        rect = fitz.Rect(x, y_base + 180 - bar_h, x + 15, y_base + 180)
        page.draw_rect(rect, color=theme["accent"], fill=theme["accent"])
        if h % 6 == 0:
            draw_text_safe(page, x, y_base + 195, f"{h}h", 10, theme["text"])

def _draw_slide_landscape(doc, stats, title, subtitle):
    page, theme = _get_page(doc, 0)
    draw_text_safe(page, 0, 80, title, 30, theme["accent"], "helvbo", align=1)
    draw_text_safe(page, 0, 130, subtitle, 16, theme["text"], "helv", align=1)

    sorted_projects = sorted(stats["deep_folder_activity"].items(), key=lambda x: x[1]["count"], reverse=True)[:7]
    if not sorted_projects:
        draw_text_safe(page, 0, 300, "No significant projects found.", 14, theme["text"], align=1)
        return

    start_y, start_x, chart_width, bar_height, gap = 200, 50, 450, 30, 50
    max_count = sorted_projects[0][1]["count"] if sorted_projects else 1

    for i, (proj_name, data) in enumerate(sorted_projects):
        y = start_y + (i * (bar_height + gap))
        display_name = proj_name if len(proj_name) <= 55 else "..." + proj_name[-52:]
        draw_text_safe(page, start_x, y - 12, f"{i+1}. {display_name}", 10, theme["text"], "helvbo")

        bar_w = (data["count"] / max_count) * chart_width
        rect = fitz.Rect(start_x, y, start_x + bar_w, y + bar_height)
        color = theme["accent"] if i == 0 else theme["sec"] 
        
        if bar_w > 0: page.draw_rect(rect, color=color, fill=color)
        draw_text_safe(page, start_x + bar_w + 10, y + 8, f"{data['count']}", 12, theme["text"])

def _get_month_name(year, week_num):
    try:
        return datetime.datetime.strptime(f"{year}-W{week_num}-1", "%Y-W%W-%w").strftime("%B")
    except: return "Unknown"

def _generate_ai_sentence(generator, keywords):
    if not generator or not keywords: return None
    kw_str = ", ".join(keywords)
    try:
        res = generator(f"Describe the relationship between these topics in a single, direct, factual sentence without introductory phrases: {kw_str}.", max_length=40, do_sample=True, temperature=0.7)
        return res[0]['generated_text']
    except Exception: return None

def _draw_slide_aura_weekly(doc, weekly_data, title, year):
    page, theme = _get_page(doc, 0)
    draw_text_safe(page, 0, 30, title, 30, theme["accent"], "helvbo", align=1)
    
    if not weekly_data: return draw_text_safe(page, 0, 300, "Not enough data.", 14, theme["text"], align=1)

    generator, doc_freq = _get_ai_generator(), collections.Counter()
    for wk in weekly_data: doc_freq.update(set(weekly_data[wk]["words"].elements()))
    
    grouped_data = collections.defaultdict(list)
    total_weeks = len(weekly_data)
    
    for wk in sorted(weekly_data.keys()):
        words_counter = weekly_data[wk]["words"]
        if not words_counter: continue

        scored_words = sorted([(count * (math.log(total_weeks / (1 + doc_freq[w])) + 1), w) for w, count in words_counter.items()], reverse=True)
        top_words = [w for s, w in scored_words[:4]] 
        
        proof_file = ""
        if top_words:
            best_word = top_words[0]
            max_occur = 0
            for fname, fwords in weekly_data[wk]["files"].items():
                if (cnt := fwords.count(best_word)) > max_occur:
                    max_occur, proof_file = cnt, fname
        
        grouped_data[_get_month_name(year, wk)].append((wk, top_words, proof_file))

    month_order = {m: i for i, m in enumerate(calendar.month_name) if m}
    current_y = 100
    
    for month in sorted(grouped_data.keys(), key=lambda m: month_order.get(m, 99)):
        if current_y > 720: break
        draw_text_safe(page, 50, current_y, month.upper(), 14, theme["accent"], "helvbo")
        current_y += 20
        
        for wk, top_words, proof in grouped_data[month]:
            if current_y > 750: break
            sentence = _generate_ai_sentence(generator, top_words) if generator else None
            sentence = (sentence[0].upper() + sentence[1:]) if sentence else f"{', '.join([w.title() for w in top_words[:3]])} were key themes."
            
            draw_text_safe(page, 60, current_y, f"W{wk:02}:", 10, theme["sec"], "helvbo")
            draw_text_safe(page, 100, current_y, sentence, 10, theme["text"], "helv")
            draw_text_safe(page, 100, current_y + 11, f"Src: {proof[:10] + '...' + proof[-5:] if len(proof) > 20 else proof}", 8, (0.5, 0.5, 0.5), "helv")
            current_y += 28
        current_y += 10

def _draw_slide_heatmap(doc, year, stats):
    page = doc.new_page(width=600, height=800)
    page.draw_rect(page.rect, color=(1, 1, 1), fill=(1, 1, 1))
    page.insert_text((50, 50), "Your Consistency", fontsize=24, color=(0,0,0))
    page.insert_text((50, 80), "Daily activity map", fontsize=12, color=(0.5, 0.5, 0.5))

    daily_counts = stats["daily_counts"]
    max_daily = max(daily_counts.values()) if daily_counts else 1
    palette = [(0.95, 0.95, 0.95), (0.8, 1.0, 0.8), (0.6, 1.0, 0.6), (0.3, 0.8, 0.3), (0.1, 0.6, 0.1)]
    current_y, box_size, gap, month_gap = 120, 11, 2, 15

    for row in range(4):
        for col in range(3):
            m = row * 3 + col + 1 
            x_base = 50 + (col * (7 * (box_size + gap) + month_gap))
            page.insert_text((x_base, current_y - 5), calendar.month_abbr[m], fontsize=10, color=(0.3, 0.3, 0.3))
            
            for week_idx, week in enumerate(calendar.monthcalendar(year, m)):
                for day_idx, day in enumerate(week):
                    if day == 0: continue
                    level = min(4, 1 + int((daily_counts.get((m, day), 0) / max_daily) * 3)) if daily_counts.get((m, day), 0) > 0 else 0
                    color = palette[level]
                    r_x = x_base + (day_idx * (box_size + gap))
                    r_y = current_y + (week_idx * (box_size + gap))
                    page.draw_rect(fitz.Rect(r_x, r_y, r_x + box_size, r_y + box_size), color=color, fill=color)
        current_y += (7 * (box_size + gap)) + 40