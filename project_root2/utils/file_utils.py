import os
import sys

def get_data_path(filename):
    """Return absolute path to file in data directory."""
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base, 'data')
    os.makedirs(data_dir, exist_ok=True)
    return os.path.join(data_dir, filename)

def get_db_path(db_name):
    """Return path for database file."""
    return get_data_path(db_name)