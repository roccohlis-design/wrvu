#!/usr/bin/env python3
"""
Signed Exam → wRVU (CMS-backed) desktop app
Windows, Python 3.13, Tkinter

Features:
- Parse CMS RVU files (CSV/XLSX/ZIP) to build CPT→work RVU cache
- Normalize free-text exam lines into canonical exam types with mapped CPTs
- PowerScribe One capture with window focusing and clipboard extraction
- Look-back filtering to prevent prior-exam pollution
- Bottom toolbar showing Total wRVU, wRVU/hr, and counts
- CSV report export
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv
import re
import zipfile
import io
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Optional, Set
from collections import defaultdict
import ctypes
from ctypes import wintypes
import time

# ============================================================================
# Windows API for PowerScribe One window focusing and keyboard control
# ============================================================================

# Virtual key codes
VK_PRIOR = 0x21      # Page Up
VK_NEXT = 0x22       # Page Down
VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_MENU = 0x12       # Alt key

# Window show commands
SW_RESTORE = 9

# Constants
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002

# Load user32.dll
user32 = ctypes.windll.user32

# Window enumeration callback type
EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def _find_powerscribe_one_hwnd() -> Optional[int]:
    """Find the top-level window whose title contains 'PowerScribe One' (case-insensitive)."""
    found_hwnd = None

    def enum_callback(hwnd, lparam):
        nonlocal found_hwnd
        if user32.IsWindowVisible(hwnd):
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buff = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buff, length + 1)
                title = buff.value
                if title and "powerscribe one" in title.lower():
                    found_hwnd = hwnd
                    return False  # Stop enumeration
        return True  # Continue enumeration

    user32.EnumWindows(EnumWindowsProc(enum_callback), 0)
    return found_hwnd


def _force_foreground(hwnd: int) -> bool:
    """Force a window to the foreground using multiple techniques."""
    try:
        # Restore if minimized
        user32.ShowWindow(hwnd, SW_RESTORE)
        time.sleep(0.05)

        # Bring to top
        user32.BringWindowToTop(hwnd)
        time.sleep(0.05)

        # Set foreground
        user32.SetForegroundWindow(hwnd)
        time.sleep(0.05)

        # Minimal Alt tap to satisfy Windows focus rules
        user32.keybd_event(VK_MENU, 0, 0, 0)
        time.sleep(0.02)
        user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.03)

        return True
    except Exception:
        return False


def _key_down(vk: int):
    """Send a key down event."""
    user32.keybd_event(vk, 0, 0, 0)


def _key_up(vk: int):
    """Send a key up event."""
    user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)


def _tap(vk: int, hold_ms: int = 30):
    """Send a key press (down + up) with configurable hold time."""
    _key_down(vk)
    time.sleep(hold_ms / 1000.0)
    _key_up(vk)
    time.sleep(0.03)  # Brief pause between keys


# ============================================================================
# CMS RVU Data Parsing
# ============================================================================

def parse_cms_rvu_file(file_path: str) -> Dict[str, float]:
    """
    Parse CMS RVU file (CSV, XLSX, or ZIP containing either).
    Returns dict mapping CPT code (str) → work RVU (float).
    """
    rvu_cache = {}
    path = Path(file_path)

    # Handle ZIP files
    if path.suffix.lower() == '.zip':
        with zipfile.ZipFile(file_path, 'r') as zf:
            for name in zf.namelist():
                if name.lower().endswith(('.csv', '.xlsx')):
                    with zf.open(name) as f:
                        data = f.read()
                        if name.lower().endswith('.csv'):
                            rvu_cache.update(_parse_cms_csv(io.StringIO(data.decode('utf-8', errors='ignore'))))
                        else:
                            rvu_cache.update(_parse_cms_xlsx(io.BytesIO(data)))
                    break  # Use first valid file
        return rvu_cache

    # Handle direct CSV
    if path.suffix.lower() == '.csv':
        with open(file_path, 'r', encoding='utf-8-sig', errors='ignore') as f:
            rvu_cache = _parse_cms_csv(f)

    # Handle direct XLSX
    elif path.suffix.lower() in ['.xlsx', '.xls']:
        rvu_cache = _parse_cms_xlsx(file_path)

    return rvu_cache


def _parse_cms_csv(file_obj) -> Dict[str, float]:
    """Parse CMS CSV format looking for CPT code and work RVU columns."""
    rvu_cache = {}
    reader = csv.DictReader(file_obj)

    # Find column names (case-insensitive matching)
    if not reader.fieldnames:
        return rvu_cache

    cpt_col = None
    wrvu_col = None

    # Debug: print available columns
    print(f"DEBUG: CSV columns found: {reader.fieldnames}")

    for col in reader.fieldnames:
        col_lower = col.lower().strip()
        # Look for CPT/code column (more flexible)
        if cpt_col is None:
            if any(keyword in col_lower for keyword in ['cpt', 'hcpcs', 'code', 'proc', 'procedure']):
                cpt_col = col
                print(f"DEBUG: Using CPT column: {col}")
        # Look for work RVU column (more flexible)
        if wrvu_col is None:
            if any(keyword in col_lower for keyword in ['work rvu', 'work_rvu', 'wrvu', 'work rvus', 'rvu']):
                wrvu_col = col
                print(f"DEBUG: Using wRVU column: {col}")

    if not cpt_col or not wrvu_col:
        print(f"DEBUG: Missing columns - CPT: {cpt_col}, wRVU: {wrvu_col}")
        return rvu_cache

    for row in reader:
        try:
            cpt = str(row.get(cpt_col, '')).strip()
            wrvu_str = str(row.get(wrvu_col, '')).strip()

            # Extract numeric CPT code
            cpt_match = re.search(r'\d{4,5}', cpt)
            if cpt_match:
                cpt_code = cpt_match.group()
                wrvu = float(wrvu_str)
                rvu_cache[cpt_code] = wrvu
        except (ValueError, AttributeError):
            continue

    print(f"DEBUG: Parsed {len(rvu_cache)} CPT codes from CSV")
    return rvu_cache


def _parse_cms_xlsx(file_path) -> Dict[str, float]:
    """Parse CMS Excel format. Falls back to CSV-like parsing if pandas unavailable."""
    rvu_cache = {}

    try:
        import pandas as pd
        df = pd.read_excel(file_path)

        # Find columns
        cpt_col = None
        wrvu_col = None

        for col in df.columns:
            col_lower = str(col).lower().strip()
            if 'cpt' in col_lower or 'hcpcs' in col_lower or 'code' in col_lower:
                if not cpt_col:
                    cpt_col = col
            if 'work rvu' in col_lower or 'work_rvu' in col_lower or 'wrvu' in col_lower:
                if not wrvu_col:
                    wrvu_col = col

        if cpt_col and wrvu_col:
            for _, row in df.iterrows():
                try:
                    cpt = str(row[cpt_col]).strip()
                    cpt_match = re.search(r'\d{4,5}', cpt)
                    if cpt_match:
                        cpt_code = cpt_match.group()
                        wrvu = float(row[wrvu_col])
                        rvu_cache[cpt_code] = wrvu
                except (ValueError, AttributeError, KeyError):
                    continue

    except ImportError:
        # Fallback: try openpyxl directly
        try:
            from openpyxl import load_workbook
            wb = load_workbook(file_path, read_only=True, data_only=True)
            ws = wb.active

            # Read header row
            headers = []
            for cell in ws[1]:
                headers.append(str(cell.value or '').strip())

            cpt_idx = None
            wrvu_idx = None

            for idx, h in enumerate(headers):
                h_lower = h.lower()
                if 'cpt' in h_lower or 'hcpcs' in h_lower or 'code' in h_lower:
                    if cpt_idx is None:
                        cpt_idx = idx
                if 'work rvu' in h_lower or 'work_rvu' in h_lower or 'wrvu' in h_lower:
                    if wrvu_idx is None:
                        wrvu_idx = idx

            if cpt_idx is not None and wrvu_idx is not None:
                for row in ws.iter_rows(min_row=2, values_only=True):
                    try:
                        cpt = str(row[cpt_idx] or '').strip()
                        cpt_match = re.search(r'\d{4,5}', cpt)
                        if cpt_match:
                            cpt_code = cpt_match.group()
                            wrvu = float(row[wrvu_idx])
                            rvu_cache[cpt_code] = wrvu
                    except (ValueError, AttributeError, IndexError):
                        continue

            wb.close()
        except ImportError:
            pass  # No Excel support available

    return rvu_cache


# ============================================================================
# Exam Type → CPT Code Mapping
# ============================================================================

# Canonical exam type patterns with their CPT mappings
CANONICAL_PATTERNS = [
    # CT patterns
    (r'\bCT\s+(?:OF\s+)?HEAD\b', 'CT HEAD', ['70450', '70460', '70470']),
    (r'\bCT\s+(?:OF\s+)?(?:C[-\s]?SPINE|CERVICAL\s+SPINE)\b', 'CT C-SPINE', ['72125', '72126', '72127']),
    (r'\bCT\s+(?:OF\s+)?(?:T[-\s]?SPINE|THORACIC\s+SPINE)\b', 'CT T-SPINE', ['72128', '72129', '72130']),
    (r'\bCT\s+(?:OF\s+)?(?:L[-\s]?SPINE|LUMBAR\s+SPINE)\b', 'CT L-SPINE', ['72131', '72132', '72133']),
    (r'\bCT\s+(?:OF\s+)?CHEST\b', 'CT CHEST', ['71250', '71260', '71270']),
    (r'\bCT\s+(?:OF\s+)?(?:ABD|ABDOMEN)\b', 'CT ABDOMEN', ['74150', '74160', '74170']),
    (r'\bCT\s+(?:OF\s+)?PELVIS\b', 'CT PELVIS', ['72192', '72193', '72194']),
    (r'\bCT\s+(?:OF\s+)?(?:ABD|ABDOMEN)(?:/|[\s&]+)PELVIS\b', 'CT ABDOMEN/PELVIS', ['74176', '74177', '74178']),

    # MRI patterns
    (r'\bMRI?\s+(?:OF\s+)?BRAIN\b', 'MRI BRAIN', ['70551', '70552', '70553']),
    (r'\bMRI?\s+(?:OF\s+)?HEAD\b', 'MRI HEAD', ['70551', '70552', '70553']),
    (r'\bMRI?\s+(?:OF\s+)?(?:C[-\s]?SPINE|CERVICAL\s+SPINE)\b', 'MRI C-SPINE', ['72141', '72142', '72156']),
    (r'\bMRI?\s+(?:OF\s+)?(?:T[-\s]?SPINE|THORACIC\s+SPINE)\b', 'MRI T-SPINE', ['72146', '72147', '72157']),
    (r'\bMRI?\s+(?:OF\s+)?(?:L[-\s]?SPINE|LUMBAR\s+SPINE)\b', 'MRI L-SPINE', ['72148', '72149', '72158']),
    (r'\bMRI?\s+(?:OF\s+)?(?:ABD|ABDOMEN)\b', 'MRI ABDOMEN', ['74181', '74182', '74183']),
    (r'\bMRI?\s+(?:OF\s+)?PELVIS\b', 'MRI PELVIS', ['72195', '72196', '72197']),

    # X-ray patterns
    (r'\b(?:XR|X[-\s]?RAY)\s+(?:OF\s+)?CHEST\b', 'XR CHEST', ['71045', '71046', '71047', '71048']),
    (r'\b(?:XR|X[-\s]?RAY)\s+(?:OF\s+)?(?:ABD|ABDOMEN)\b', 'XR ABDOMEN', ['74018', '74019', '74021']),
    (r'\b(?:XR|X[-\s]?RAY)\s+(?:OF\s+)?PELVIS\b', 'XR PELVIS', ['72170', '72190']),

    # Ultrasound patterns - INCLUDING US APPENDIX
    (r'\bUS\s+(?:OF\s+)?APPENDIX\b', 'US ABDOMEN LIMITED', ['76705']),  # Specific mapping for US APPENDIX
    (r'\bUS\s+(?:OF\s+)?(?:ABD|ABDOMEN)(?:\s+(?:LTD|LIMITED))?\b', 'US ABDOMEN LIMITED', ['76705']),
    (r'\bUS\s+(?:OF\s+)?(?:ABD|ABDOMEN)(?:\s+(?:COMPLETE|COMP))?\b', 'US ABDOMEN COMPLETE', ['76700']),
    (r'\bUS\s+(?:OF\s+)?PELVIS\b', 'US PELVIS', ['76856', '76857']),
    (r'\bUS\s+(?:OF\s+)?(?:RETROPERITONEAL|RETRO)\b', 'US RETROPERITONEAL', ['76770', '76775']),
]


def normalize_exam_line(line: str) -> Tuple[str, List[str]]:
    """
    Normalize a free-text exam line to canonical exam type and CPT codes.
    Returns (canonical_name, [cpt_codes]).
    """
    line_upper = line.upper().strip()

    # Try canonical patterns first
    for pattern, canonical, cpts in CANONICAL_PATTERNS:
        if re.search(pattern, line_upper):
            return canonical, cpts

    # Fallback: return original line as canonical with empty CPT list
    return line.strip(), []


def infer_cpt_from_description(exam_text: str, cms_cache: Dict[str, float]) -> List[str]:
    """
    Inference fallback: search CMS descriptions for matching CPTs.
    This is a simplified version - real implementation would search CMS description text.
    For now, return empty list as we rely on canonical patterns.
    """
    # In a full implementation, this would search through CMS procedure descriptions
    # For this version, we rely on canonical patterns and return empty
    return []


# ============================================================================
# PowerScribe Capture Functions
# ============================================================================

def _detect_delimiter(text: str) -> str:
    """Detect whether the clipboard data is tab or comma delimited."""
    lines = text.strip().split('\n')[:5]  # Check first 5 lines

    tab_count = sum(line.count('\t') for line in lines)
    comma_count = sum(line.count(',') for line in lines)

    return '\t' if tab_count > comma_count else ','


def _extract_pairs_from_tsv(text: str) -> List[Tuple[str, str]]:
    """
    Extract (Exam Description, Modified) pairs from TSV/CSV clipboard data.
    Returns list of (exam_description_str, modified_str) tuples.
    """
    pairs = []
    delimiter = _detect_delimiter(text)

    lines = text.strip().split('\n')
    if not lines:
        return pairs

    # Try to find header row
    reader = csv.DictReader(io.StringIO(text), delimiter=delimiter)

    exam_desc_col = None
    modified_col = None

    if reader.fieldnames:
        for col in reader.fieldnames:
            col_lower = col.lower().strip()
            # Look for exam description/type column (prioritize description over date)
            if exam_desc_col is None:
                if 'description' in col_lower or 'exam type' in col_lower or 'procedure' in col_lower:
                    exam_desc_col = col
                elif col_lower in ['exam', 'exams', 'study', 'modality']:
                    exam_desc_col = col
            # Look for modified time column
            if 'modified' in col_lower or 'mod time' in col_lower or 'last modified' in col_lower:
                modified_col = col

        if exam_desc_col and modified_col:
            for row in reader:
                exam_desc = row.get(exam_desc_col, '').strip()
                modified = row.get(modified_col, '').strip()
                if exam_desc and modified:
                    pairs.append((exam_desc, modified))
            return pairs

    # Fallback: regex extraction (first text field, last datetime)
    return _extract_pairs_by_regex(text)


def _extract_pairs_by_regex(text: str) -> List[Tuple[str, str]]:
    """
    Fallback: extract datetime pairs using regex.
    Looks for two datetime stamps per line and uses (first, last).
    """
    pairs = []

    # Match common datetime formats: M/D/YYYY H:MM:SS AM/PM, M/D/YY H:MM AM/PM, etc.
    datetime_pattern = r'\d{1,2}/\d{1,2}/\d{2,4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?'

    for line in text.split('\n'):
        matches = re.findall(datetime_pattern, line, re.IGNORECASE)
        if len(matches) >= 2:
            # Use first and last
            pairs.append((matches[0].strip(), matches[-1].strip()))
        elif len(matches) == 1:
            # Use same datetime for both
            pairs.append((matches[0].strip(), matches[0].strip()))

    return pairs


def parse_datetime_flexible(dt_str: str) -> Optional[datetime]:
    """Parse common datetime formats flexibly."""
    dt_str = dt_str.strip()

    # Common formats
    formats = [
        '%m/%d/%Y %I:%M:%S %p',  # 11/3/2025 6:36:49 PM
        '%m/%d/%Y %I:%M %p',     # 11/3/2025 6:36 PM
        '%m/%d/%y %I:%M:%S %p',  # 11/3/25 6:36:49 PM
        '%m/%d/%y %I:%M %p',     # 11/3/25 6:36 PM
        '%m/%d/%Y %H:%M:%S',     # 11/3/2025 18:36:49
        '%m/%d/%Y %H:%M',        # 11/3/2025 18:36
        '%m/%d/%y %H:%M:%S',     # 11/3/25 18:36:49
        '%m/%d/%y %H:%M',        # 11/3/25 18:36
        '%Y-%m-%d %H:%M:%S',     # 2025-11-03 18:36:49
        '%Y-%m-%d %H:%M',        # 2025-11-03 18:36
    ]

    for fmt in formats:
        try:
            return datetime.strptime(dt_str, fmt)
        except ValueError:
            continue

    return None


# ============================================================================
# Main Application
# ============================================================================

class WRVUApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Signed Exam → wRVU (CMS-backed)")
        self.root.geometry("900x700")

        # Data
        self.rvu_cache: Dict[str, float] = {}
        self.seen_pairs: Set[Tuple[str, str]] = set()  # For deduplication
        self.captured_modified_times: List[datetime] = []  # For wRVU/hr calculation

        # Build UI
        self._build_ui()

    def _build_ui(self):
        """Build the complete UI."""

        # Top toolbar
        top_frame = ttk.Frame(self.root, padding=5)
        top_frame.pack(fill=tk.X)

        ttk.Button(top_frame, text="Load CMS RVU file", command=self.load_cms_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(top_frame, text="Export CSV Report", command=self.export_csv).pack(side=tk.LEFT, padx=2)

        self.cache_label = ttk.Label(top_frame, text="RVU cache entries: 0")
        self.cache_label.pack(side=tk.LEFT, padx=10)

        # Main text area for captured content
        text_frame = ttk.LabelFrame(self.root, text="Captured Exams (Exam Type — Modified)", padding=5)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.text_area = scrolledtext.ScrolledText(text_frame, height=20, width=80, wrap=tk.WORD)
        self.text_area.pack(fill=tk.BOTH, expand=True)

        # Action toolbar
        action_frame = ttk.Frame(self.root, padding=5)
        action_frame.pack(fill=tk.X)

        ttk.Button(action_frame, text="Paste Clipboard", command=self.paste_clipboard).pack(side=tk.LEFT, padx=2)
        ttk.Button(action_frame, text="Capture from PowerScribe One",
                  command=self.capture_from_powerscribe_one).pack(side=tk.LEFT, padx=2)
        ttk.Button(action_frame, text="Calculate wRVU (unique exam types)",
                  command=self.calculate_wrvu).pack(side=tk.LEFT, padx=2)

        # Look-back control
        ttk.Label(action_frame, text="Max look-back (days):").pack(side=tk.LEFT, padx=(10, 2))
        self.lookback_var = tk.StringVar(value="7")
        lookback_entry = ttk.Entry(action_frame, textvariable=self.lookback_var, width=5)
        lookback_entry.pack(side=tk.LEFT)

        # Bottom toolbar (replaces summary pane)
        bottom_frame = ttk.Frame(self.root, padding=5)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.total_wrvu_label = ttk.Label(bottom_frame, text="Total wRVU: —", font=('TkDefaultFont', 9, 'bold'))
        self.total_wrvu_label.pack(side=tk.LEFT, padx=5)

        self.wrvu_per_hour_label = ttk.Label(bottom_frame, text="wRVU/hr: —", font=('TkDefaultFont', 9, 'bold'))
        self.wrvu_per_hour_label.pack(side=tk.LEFT, padx=5)

        self.lines_label = ttk.Label(bottom_frame, text="Lines: 0")
        self.lines_label.pack(side=tk.LEFT, padx=5)

        self.types_label = ttk.Label(bottom_frame, text="Recognized types: 0")
        self.types_label.pack(side=tk.LEFT, padx=5)

    def load_cms_file(self):
        """Load CMS RVU file (CSV/XLSX/ZIP)."""
        try:
            file_path = filedialog.askopenfilename(
                title="Select CMS RVU File",
                filetypes=[
                    ("All supported", "*.csv;*.xlsx;*.xls;*.zip"),
                    ("CSV files", "*.csv"),
                    ("Excel files", "*.xlsx;*.xls"),
                    ("ZIP files", "*.zip"),
                    ("All files", "*.*")
                ]
            )

            if not file_path:
                return

            # Show loading message
            self.cache_label.config(text="Loading CMS file...")
            self.root.update()

            self.rvu_cache = parse_cms_rvu_file(file_path)
            self.cache_label.config(text=f"RVU cache entries: {len(self.rvu_cache)}")

            if len(self.rvu_cache) == 0:
                messagebox.showwarning("No Data",
                    f"No CPT codes found in file.\n\n"
                    f"File: {file_path}\n\n"
                    f"Make sure the file contains columns with:\n"
                    f"- CPT/HCPCS codes\n"
                    f"- Work RVU values")
            else:
                # Show sample CPT codes for verification
                sample = list(self.rvu_cache.items())[:5]
                sample_str = "\n".join([f"CPT {cpt}: {wrvu:.2f}" for cpt, wrvu in sample])
                messagebox.showinfo("Success",
                    f"Loaded {len(self.rvu_cache)} CPT codes from CMS file.\n\n"
                    f"Sample entries:\n{sample_str}")
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("Error",
                f"Failed to load CMS file:\n\n{str(e)}\n\n"
                f"Details:\n{error_details[:500]}")
            self.cache_label.config(text="RVU cache entries: 0 (load failed)")

    def paste_clipboard(self):
        """Paste clipboard content directly into text area."""
        try:
            clip_text = self.root.clipboard_get()
            self.text_area.insert(tk.END, clip_text + "\n")
        except tk.TclError:
            messagebox.showwarning("Clipboard Empty", "No text in clipboard.")

    def capture_from_powerscribe_one(self):
        """
        Capture from PowerScribe One:
        1. Find and focus PowerScribe One window
        2. Send PageUp, Shift+PageDown, Ctrl+C
        3. Extract Exam Description and Modified columns
        4. Append unique pairs to text area
        """
        try:
            # Find PowerScribe One window
            hwnd = _find_powerscribe_one_hwnd()
            if not hwnd:
                messagebox.showerror("Window Not Found",
                                    "Could not find PowerScribe One window.\nMake sure it's running.")
                return

            # Focus the window
            if not _force_foreground(hwnd):
                messagebox.showerror("Focus Failed",
                                    "Could not bring PowerScribe One to foreground.")
                return

            time.sleep(0.1)  # Let window settle

            # Send key sequence: PageUp
            _tap(VK_PRIOR, 35)
            time.sleep(0.05)

            # Shift+PageDown (select all from top to one page down)
            _key_down(VK_SHIFT)
            time.sleep(0.03)
            _tap(VK_NEXT, 35)
            _key_up(VK_SHIFT)
            time.sleep(0.05)

            # Ctrl+C (copy)
            _key_down(VK_CONTROL)
            time.sleep(0.03)
            _tap(0x43, 35)  # 'C' key
            _key_up(VK_CONTROL)

            # Wait for clipboard to populate
            time.sleep(0.15)

            # Read clipboard
            try:
                clip_text = self.root.clipboard_get()
            except tk.TclError:
                messagebox.showwarning("Clipboard Empty",
                                      "Clipboard is empty after capture attempt.")
                return

            if not clip_text.strip():
                messagebox.showwarning("No Data", "Captured clipboard is empty.")
                return

            # Extract pairs (with debug info)
            # First, show what we captured for debugging
            lines_preview = clip_text[:500] if len(clip_text) > 500 else clip_text
            print(f"DEBUG: Clipboard content preview:\n{lines_preview}\n")

            pairs = _extract_pairs_from_tsv(clip_text)

            if not pairs:
                # Show first few lines to help debug
                first_lines = '\n'.join(clip_text.split('\n')[:5])
                messagebox.showwarning("No Data",
                                      f"Could not extract exam description/modified pairs from clipboard.\n\n"
                                      f"First few lines captured:\n{first_lines}\n\n"
                                      f"Make sure the table has exam description and modified time columns.")
                return

            # Apply look-back filter
            try:
                lookback_days = int(self.lookback_var.get())
            except ValueError:
                lookback_days = 7

            cutoff = datetime.now() - timedelta(days=lookback_days)

            new_count = 0
            filtered_count = 0

            for exam_date_str, modified_str in pairs:
                # Check if already seen
                pair_key = (exam_date_str, modified_str)
                if pair_key in self.seen_pairs:
                    continue

                # Parse modified datetime
                modified_dt = parse_datetime_flexible(modified_str)

                # Apply look-back filter
                if modified_dt and modified_dt < cutoff:
                    filtered_count += 1
                    continue

                # Add to text area
                self.text_area.insert(tk.END, f"{exam_date_str} — {modified_str}\n")
                self.seen_pairs.add(pair_key)

                # Store modified datetime for wRVU/hr calculation
                if modified_dt:
                    self.captured_modified_times.append(modified_dt)

                new_count += 1

            msg = f"Captured {new_count} new exam(s)."
            if filtered_count > 0:
                msg += f"\n{filtered_count} older exam(s) filtered out (beyond {lookback_days}-day look-back)."

            messagebox.showinfo("Capture Complete", msg)

        except Exception as e:
            messagebox.showerror("Capture Error", f"An error occurred during capture:\n{e}")

    def calculate_wrvu(self):
        """Calculate wRVU from captured exam lines."""
        if not self.rvu_cache:
            messagebox.showwarning("No RVU Cache",
                                  "Please load a CMS RVU file first.")
            return

        text = self.text_area.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("No Data", "No exam data to calculate.")
            return

        lines = [l.strip() for l in text.split('\n') if l.strip()]

        # Extract exam type from each line (before the em-dash)
        exam_types_seen = set()
        exam_type_counts = defaultdict(int)
        exam_type_details = []  # For debugging
        total_wrvu = 0.0
        unmatched_count = 0

        for line in lines:
            # Split on em-dash to get exam type
            if '—' in line:
                exam_part = line.split('—')[0].strip()
            else:
                exam_part = line

            # Normalize to canonical exam type and CPTs
            canonical, cpts = normalize_exam_line(exam_part)

            # Use canonical as the unique key
            if canonical in exam_types_seen:
                continue  # Already counted this type

            exam_types_seen.add(canonical)
            exam_type_counts[canonical] += 1

            # Calculate wRVU for this exam type
            wrvu = 0.0
            cpt_used = ""
            if cpts:
                # Use first CPT code as primary (could be enhanced to pick best)
                primary_cpt = cpts[0]
                cpt_used = primary_cpt
                wrvu = self.rvu_cache.get(primary_cpt, 0.0)
                total_wrvu += wrvu
            else:
                # Try inference fallback
                inferred_cpts = infer_cpt_from_description(canonical, self.rvu_cache)
                if inferred_cpts:
                    cpt_used = inferred_cpts[0]
                    wrvu = self.rvu_cache.get(inferred_cpts[0], 0.0)
                    total_wrvu += wrvu
                else:
                    unmatched_count += 1

            # Store details for debugging
            exam_type_details.append(f"{canonical} → CPT {cpt_used} = {wrvu:.2f} wRVU")

        # Calculate wRVU per hour
        wrvu_per_hour = self.compute_wrvu_per_hour(self.captured_modified_times)

        # Update bottom toolbar
        self.total_wrvu_label.config(text=f"Total wRVU: {total_wrvu:.2f}")

        if wrvu_per_hour is not None:
            self.wrvu_per_hour_label.config(text=f"wRVU/hr: {wrvu_per_hour:.2f}")
        else:
            self.wrvu_per_hour_label.config(text="wRVU/hr: —")

        self.lines_label.config(text=f"Lines: {len(lines)}")
        self.types_label.config(text=f"Recognized types: {len(exam_types_seen) - unmatched_count}")

        # Show detailed results
        result_msg = f"Total wRVU: {total_wrvu:.2f}\n"
        if wrvu_per_hour is not None:
            result_msg += f"wRVU/hr: {wrvu_per_hour:.2f}\n"
        result_msg += f"Lines: {len(lines)}\n"
        result_msg += f"Unique exam types: {len(exam_types_seen)}\n"
        result_msg += f"Recognized: {len(exam_types_seen) - unmatched_count}\n"
        result_msg += f"Unmatched: {unmatched_count}\n\n"
        result_msg += "Breakdown:\n" + "\n".join(exam_type_details)

        messagebox.showinfo("Calculation Complete", result_msg)

    def compute_wrvu_per_hour(self, modified_datetimes: List[datetime]) -> Optional[float]:
        """
        Compute wRVU per hour based on time span of modified datetimes.
        Returns None if insufficient data.
        Clamps to minimum 0.25 hours if span is < 15 minutes.
        """
        if not modified_datetimes or not self.total_wrvu_label.cget("text").startswith("Total wRVU: "):
            return None

        # Get total wRVU from label
        try:
            wrvu_text = self.total_wrvu_label.cget("text")
            total_wrvu = float(wrvu_text.split(": ")[1])
        except (ValueError, IndexError):
            return None

        if len(modified_datetimes) < 1:
            return None

        earliest = min(modified_datetimes)
        latest = max(modified_datetimes)

        time_span = latest - earliest
        hours = time_span.total_seconds() / 3600.0

        # Clamp to minimum 0.25 hours (15 minutes)
        if hours < 0.25:
            hours = 0.25

        return total_wrvu / hours if hours > 0 else None

    def export_csv(self):
        """Export calculation results to CSV."""
        if not self.rvu_cache:
            messagebox.showwarning("No RVU Cache",
                                  "Please load a CMS RVU file first.")
            return

        text = self.text_area.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("No Data", "No exam data to export.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Export CSV Report",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            lines = [l.strip() for l in text.split('\n') if l.strip()]

            # Collect unique exam types with details
            exam_type_details = {}

            for line in lines:
                # Split on em-dash to get exam type
                if '—' in line:
                    exam_part = line.split('—')[0].strip()
                else:
                    exam_part = line

                # Normalize to canonical exam type and CPTs
                canonical, cpts = normalize_exam_line(exam_part)

                if canonical in exam_type_details:
                    continue

                # Get wRVU
                wrvu = 0.0
                cpt_used = ""
                if cpts:
                    cpt_used = cpts[0]
                    wrvu = self.rvu_cache.get(cpt_used, 0.0)
                else:
                    inferred = infer_cpt_from_description(canonical, self.rvu_cache)
                    if inferred:
                        cpt_used = inferred[0]
                        wrvu = self.rvu_cache.get(cpt_used, 0.0)

                exam_type_details[canonical] = {
                    'cpt': cpt_used,
                    'wrvu': wrvu
                }

            # Write CSV
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Exam Type', 'CPT Code', 'Work RVU'])

                for exam_type, details in sorted(exam_type_details.items()):
                    writer.writerow([exam_type, details['cpt'], f"{details['wrvu']:.2f}"])

                # Summary rows
                writer.writerow([])
                writer.writerow(['Summary', '', ''])
                writer.writerow(['Total wRVU', '', self.total_wrvu_label.cget("text").split(": ")[1]])
                writer.writerow(['wRVU/hr', '', self.wrvu_per_hour_label.cget("text").split(": ")[1]])
                writer.writerow(['Total Lines', '', self.lines_label.cget("text").split(": ")[1]])
                writer.writerow(['Recognized Types', '', self.types_label.cget("text").split(": ")[1]])

            messagebox.showinfo("Export Complete", f"Report saved to:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export CSV:\n{e}")


# ============================================================================
# Entry Point
# ============================================================================

def main():
    root = tk.Tk()
    app = WRVUApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
