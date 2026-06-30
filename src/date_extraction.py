# --- date_extraction.py ---
import os
import re
from datetime import datetime

def parse_token_date(token):
    """
    Accepts strings containing YYYYMMDD or YYYY 
    anywhere in the token and returns (year, month, day), or None if it doesn’t match.
    """
    # Look for exactly 8 digits (YYYYMMDD) avoiding surrounding digits
    m = re.search(r'(?<!\d)(\d{8})(?!\d)', token)
    if m:
        s = m.group(1)
        return (int(s[0:4]), int(s[4:6]), int(s[6:8]))
    
    # Look for exactly 4 digits (YYYY)
    m = re.search(r'(?<!\d)(\d{4})(?!\d)', token)
    if m:
        s = m.group(1)
        return (int(s), 1, 1)
    
    return None

def extract_year_key(path):
    """
    Generates a sort key from a file path based on a date-like prefix.
    The primary key is the parsed date, and the secondary key is the filename.
    """
    name = os.path.basename(path)
    bare = os.path.splitext(name)[0]
    tok = bare.split()[0] if bare else ""
    parsed = parse_token_date(tok)
    
    if parsed:
        y, mo, da = parsed
        date_key = (y, mo, da)
    else:
        # -inf ensures non-dated files sink to the bottom of descending sorts
        date_key = (float('-inf'), 0, 0)
        
    return (date_key, name.lower())

def get_datetime_from_path(path):
    """
    Helper for documentinfo to get a valid datetime object if possible.
    """
    name = os.path.basename(path)
    bare = os.path.splitext(name)[0]
    tok = bare.split()[0] if bare else ""
    parsed = parse_token_date(tok)
    
    if parsed:
        y, m, d = parsed
        # Validate that the year is positive and dates are valid for a datetime object
        if y > 0 and 1 <= m <= 12 and 1 <= d <= 31:
            try:
                return datetime(y, m, d)
            except ValueError:
                pass
    return None