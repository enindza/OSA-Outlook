"""
Outlook EML Saver — hotkey mod
- Kliknes na mejl u Outlooku
- Pritisnes Ctrl+Shift+S → snimi kao .eml + premesti u Inbox2
- Pritisnes Ctrl+Shift+Q → ugasi skriptu
"""
import sys, io, time, os, re, csv, threading, shutil, email, ctypes
import configparser, pathlib
from email.header import decode_header as eml_decode_header
import tkinter as tk

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# ──────────────────────────────────────────────
# Konfiguracija (config.ini)
# ──────────────────────────────────────────────

_SCRIPT_DIR = pathlib.Path(__file__).parent
_CFG_PATH   = _SCRIPT_DIR / "config.ini"

if not _CFG_PATH.exists():
    sys.exit(f"GRESKA: config.ini nije pronadjen u {_SCRIPT_DIR}\n"
             f"Kopiraj config.ini.example u config.ini i podesiti putanje.")

_cfg = configparser.ConfigParser(inline_comment_prefixes=("#",), interpolation=None)
_cfg.read(_CFG_PATH, encoding="utf-8")

def _path(section, key):
    """Cita putanju iz config-a i razvija %env% varijable."""
    return os.path.expandvars(_cfg.get(section, key))

def _relpath(section, key):
    """Cita putanju — ako je relativna, resolvuje od foldera skripte."""
    p = _path(section, key)
    return p if pathlib.Path(p).is_absolute() else str(_SCRIPT_DIR / p)

TERMINAL_LOG = str(_SCRIPT_DIR / "terminal.log")

class Tee:
    """Prosleđuje upis i na terminal i u fajl."""
    def __init__(self, stream, filepath):
        self._stream = stream
        self._file   = open(filepath, "a", encoding="utf-8", errors="replace")

    def write(self, data):
        self._stream.write(data)
        self._stream.flush()
        self._file.write(data)
        self._file.flush()

    def flush(self):
        self._stream.flush()
        self._file.flush()

    def close_log(self):
        """Zatvara samo log fajl, ne i terminal stream."""
        try:
            self._file.close()
        except Exception:
            pass

    def __getattr__(self, name):
        """Prosledi sve ostale atribute na originalni stream."""
        return getattr(self._stream, name)

_tee_out = Tee(sys.stdout, TERMINAL_LOG)
_tee_err = Tee(sys.stderr, TERMINAL_LOG)
sys.stdout = _tee_out
sys.stderr = _tee_err

import pyautogui
import pyperclip
import keyboard
from datetime import datetime
from pywinauto import Desktop

pyautogui.FAILSAFE = False

SAVE_FOLDER      = _path("paths",    "save_folder")
LOG_FILE         = _relpath("paths", "log_file")
HOTKEY_SAVE      = _cfg.get("hotkeys",   "save")
DRY_RUN          = _cfg.getboolean("behaviour", "dry_run")
DOWNLOADS_FOLDER = _path("paths",    "downloads_folder")

# Kasnjenja ucitana iz config.ini — koriste se kroz celu skriptu
_D_CTX   = _cfg.getfloat("delays", "context_menu_open")  # Shift+F10
_D_FILE  = _cfg.getfloat("delays", "file_menu_open")      # Alt+F
_D_SUB   = _cfg.getfloat("delays", "submenu_open")        # Right arrow
_D_SA    = _cfg.getfloat("delays", "save_as_open")        # Enter → dialog
_D_ALTN  = _cfg.getfloat("delays", "alt_n")               # Alt+N
_D_NAV   = _cfg.getfloat("delays", "key_nav")             # navigacija menijem

# Retry parametri
MAX_DIALOG_RETRY  = _cfg.getint("retries", "max_dialog_retry")
DIALOG_RETRY_WAIT = _cfg.getint("retries", "dialog_retry_wait")
MAX_GROUP_RETRY   = _cfg.getint("retries", "max_group_retry")

# Flag: posle snimanja, sledeci hotkey poziv ide u "move" mod
_pending_move    = False
_pending_move_lk = threading.Lock()
_pending_subject = ""   # cuva subject izmedju save i move faze
_pending_file    = ""   # cuva filename izmedju save i move faze

# Odredisni folder za premestanje (unosi se pri startu)
DEST_FOLDER = _cfg.get("behaviour", "default_dest_folder")

# Batch brojaci
MAX_BATCH = _cfg.getint("behaviour", "max_batch")
_batch_remaining = 0   # koliko mejlova jos treba obraditi
_batch_done      = 0   # koliko je uspesno obradjeno u tekucem batchu
_stop_requested  = False  # True = zavrsi tekuci mejl pa stani


# ──────────────────────────────────────────────
# Toast notifikacija
# ──────────────────────────────────────────────

def toast(poruka, boja="#222222", trajanje=2.5):
    def _show():
        root = tk.Tk()
        root.overrideredirect(True)
        root.attributes("-topmost", True)
        root.attributes("-alpha", 0.92)
        lbl = tk.Label(root, text=poruka, bg=boja, fg="white",
                       font=("Segoe UI", 11), padx=18, pady=12,
                       wraplength=380, justify="left")
        lbl.pack()
        root.update_idletasks()
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        w  = root.winfo_width()
        h  = root.winfo_height()
        root.geometry(f"+{sw - w - 30}+{sh - h - 60}")
        root.after(int(trajanje * 1000), root.destroy)
        root.mainloop()
    threading.Thread(target=_show, daemon=True).start()


# ──────────────────────────────────────────────
# Log
# ──────────────────────────────────────────────

def init_log():
    exists = os.path.isfile(LOG_FILE)
    f = open(LOG_FILE, "a", newline="", encoding="utf-8")
    w = csv.writer(f)
    if not exists:
        w.writerow(["timestamp", "predmet", "ime_fajla", "status", "napomena"])
    return f, w

_log_file = None
_log_writer = None

def log_entry(predmet, filename, status, note=""):
    if _log_writer:
        _log_writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            predmet, filename, status, note
        ])
        _log_file.flush()


# ──────────────────────────────────────────────
# Outlook
# ──────────────────────────────────────────────

def get_outlook_win():
    try:
        result = []
        for w in Desktop(backend="uia").windows():
            try:
                if "Outlook Host" in (w.class_name() or ""):
                    result.append(w)
            except Exception:
                pass
        return result[0] if result else None
    except Exception:
        return None


def get_outlook_win_retry(max_pokusaja=5, pauza=10):
    """Pokusava da pronadje Outlook prozor do max_pokusaja puta, sa pauzom izmedju."""
    for i in range(1, max_pokusaja + 1):
        win = get_outlook_win()
        if win is not None:
            return win
        if i < max_pokusaja:
            print(f"  Outlook nije pronadjen (pokusaj {i}/{max_pokusaja}) — cekam {pauza}s...")
            toast(f"Outlook nije pronadjen ({i}/{max_pokusaja})\nPonavljam za {pauza}s...",
                  boja="#b05000", trajanje=pauza)
            time.sleep(pauza)
        else:
            print(f"  Outlook nije pronadjen ni posle {max_pokusaja} pokusaja — odustajem.")
    return None


def sanitize_filename(text, max_len=100):
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", text)
    return text.strip(" .")[:max_len]


def unique_filename(base_name):
    """Ako fajl vec postoji u SAVE_FOLDER, dodaje _1, _2, ... dok ne nadje slobodno ime.
    Ako su svi numericki slotovi zauzeti, koristi vremenski sufiks."""
    if not os.path.exists(os.path.join(SAVE_FOLDER, base_name)):
        return base_name
    stem = base_name[:-4]  # bez .eml
    for i in range(1, 1000):
        candidate = f"{stem}_{i}.eml"
        if not os.path.exists(os.path.join(SAVE_FOLDER, candidate)):
            return candidate
    # Svi slotovi 1-999 zauzeti — koristi timestamp
    candidate = f"{stem}_{datetime.now().strftime('%H%M%S')}.eml"
    return candidate


# ──────────────────────────────────────────────
# Detekcija tipa selektovanog itema
# ──────────────────────────────────────────────

def is_meeting_item(win):
    """
    Vraca (True, naziv) ako je fokusiran ListItem meeting/Teams pozivnica,
    (False, naziv) ako je obican mejl,
    (None, "") ako fokusiran item nije pronadjen.

    Outlook u UIA-u prefixuje naziv meeting itema sa 'Meeting' ili
    'Meeting start time', dok obicni mejlovi pocinju imenom posiljaoca
    ili 'Has attachments'.
    """
    try:
        for elem in win.descendants():
            try:
                ei = elem.element_info
                if ei.control_type != "ListItem":
                    continue
                name = (ei.name or "").strip()
                if not name:
                    continue
                try:
                    focused = bool(ei._element.CurrentHasKeyboardFocus)
                except Exception:
                    focused = False
                if focused:
                    return name.lower().startswith("meeting"), name
            except Exception:
                pass
    except Exception:
        pass
    return None, ""


# ──────────────────────────────────────────────
# Navigacija na gornji mejl u listi
# ──────────────────────────────────────────────

def navigate_to_top(win):
    """Fokus na listu mejlova: Esc×2 → Ctrl+1 → 5×F6."""
    win.set_focus()
    time.sleep(0.20)
    print("  [TOP] Esc → Ctrl+1 → Esc → Alt → 3×F6...")
    pyautogui.press("escape");       time.sleep(0.05)
    pyautogui.hotkey("ctrl", "1");   time.sleep(0.25)
    pyautogui.press("escape");       time.sleep(0.10)
    pyautogui.press("alt");          time.sleep(0.10)
    pyautogui.press("f6");           time.sleep(0.10)
    pyautogui.press("f6");           time.sleep(0.10)
    pyautogui.press("f6");           time.sleep(0.10)
    pyautogui.press("home");         time.sleep(0.05)
    pyautogui.press("down");         time.sleep(0.05)
    pyautogui.press("up");           time.sleep(0.05)


# ──────────────────────────────────────────────
# Save As via context meni (Teams pozivnice)
# ──────────────────────────────────────────────

def open_saveas_via_contextmenu(win):
    """
    Otvara Save As dialog za Teams meeting putem Shift+F10 context menija:
    Shift+F10 → End → Up×3 (Save as) → Right (podmeni) → Down (Save as EML) → Enter
    """
    win.set_focus()
    time.sleep(0.20)
    print("  [MEETING] Shift+F10...")
    pyautogui.hotkey("shift", "f10"); time.sleep(_D_CTX)
    print("  [MEETING] End → Up×3 → Save as...")
    pyautogui.press("end");           time.sleep(0.15)
    for _ in range(3):
        pyautogui.press("up"); time.sleep(_D_NAV)
    print("  [MEETING] Right → podmeni → Enter (Save as EML)...")
    pyautogui.press("right"); time.sleep(_D_SUB)
    pyautogui.press("enter"); time.sleep(_D_SA)


# ──────────────────────────────────────────────
# Download detekcija
# ──────────────────────────────────────────────

def snapshot_downloads():
    """Vraca skup fajlova trenutno u Downloads folderu."""
    try:
        return set(os.listdir(DOWNLOADS_FOLDER))
    except Exception:
        return set()


def wait_for_download(before_snapshot, total_timeout=300):
    """
    Faza 1: ceka da se pojavi bilo koji novi fajl u Downloads (download poceo).
    Faza 2: ceka da taj fajl postane .eml (download zavrsen).
    Vraca set novih .eml fajlova, ili prazan set ako timeout.
    """
    toast("Outlook preuzima fajl...\nCekam.", boja="#1a5276", trajanje=15)
    deadline = time.time() + total_timeout

    # Faza 1: cekaj da se bilo sta pojavi (privremeni fajl, .crdownload, itd.)
    print(f"  [DOWNLOAD] Faza 1 — cekam pocetak downloada (do {total_timeout}s)...")
    while time.time() < deadline:
        current = snapshot_downloads()
        novi    = current - before_snapshot
        if novi:
            print(f"  [DOWNLOAD] Faza 1 — pojavio se: {novi}")
            break
        time.sleep(1)
    else:
        print(f"  [DOWNLOAD] Timeout — nista se nije pojavilo u Downloads.")
        return set()

    # Faza 2: cekaj da taj fajl postane .eml
    print(f"  [DOWNLOAD] Faza 2 — cekam da download bude zavrseno...")
    while time.time() < deadline:
        current  = snapshot_downloads()
        novi_eml = {f for f in (current - before_snapshot)
                    if f.lower().endswith(".eml")}
        if novi_eml:
            print(f"  [DOWNLOAD] Gotovo: {novi_eml}")
            return novi_eml
        time.sleep(2)

    print(f"  [DOWNLOAD] Timeout — .eml nije kompletiran.")
    return set()


def subject_from_eml(filepath):
    """Cita Subject header iz EML fajla."""
    try:
        with open(filepath, "rb") as f:
            msg = email.message_from_binary_file(f)
        raw = msg.get("Subject", "")
        parts = eml_decode_header(raw)
        subject = ""
        for part, enc in parts:
            if isinstance(part, bytes):
                subject += part.decode(enc or "utf-8", errors="replace")
            else:
                subject += part
        return subject.strip() or "email"
    except Exception as e:
        print(f"  [DOWNLOAD] Ne mogu da procitam Subject: {e}")
        return "email"


def handle_download_file(before_snapshot, date_str):
    """
    Kada Save As dialog ne postoji, trazi novi EML fajl u Downloads folderu,
    cita Subject, preimenuje po konvenciji i premesta u SAVE_FOLDER.
    Vraca (filename, "ok", subject) ili (None, "error:...", "")
    """
    # Cekaj da download zavrsi i dobij set gotovih fajlova
    gotovi = wait_for_download(before_snapshot)
    if not gotovi:
        return None, "error:download nije detektovan", ""

    # Preferiraj .eml fajlove, inace uzmi prvi dostupni
    eml_fajlovi = [f for f in gotovi if f.lower().endswith(".eml")]
    new_files   = eml_fajlovi if eml_fajlovi else list(gotovi)

    if not new_files:
        return None, "error:novi fajl nije pronadjen u Downloads", ""

    dl_path = os.path.join(DOWNLOADS_FOLDER, new_files[0])
    print(f"  [DOWNLOAD] Fajl preuzet: {new_files[0]}")

    # Procitaj Subject iz EML
    subject  = subject_from_eml(dl_path)
    filename = sanitize_filename(f"{date_str}_{subject}") + ".eml"
    dst_path = os.path.join(SAVE_FOLDER, filename)

    print(f"  [DOWNLOAD] Predmet : '{subject}'")
    print(f"  [DOWNLOAD] Fajl    : '{filename}'")

    try:
        shutil.move(dl_path, dst_path)
        print(f"  [DOWNLOAD] Premesten u: {dst_path}")
        return filename, "ok", subject
    except Exception as e:
        return None, f"error:premestanje nije uspelo: {e}", subject


# ──────────────────────────────────────────────
# Save dialog
# ──────────────────────────────────────────────

def snapshot_dialogs():
    """Vraca skup handle-ova svih trenutnih #32770 prozora."""
    handles = set()
    try:
        for w in Desktop(backend="win32").windows():
            try:
                if w.class_name() == "#32770":
                    handles.add(w.handle)
            except Exception:
                pass
    except Exception:
        pass
    return handles


_u32 = ctypes.windll.user32

def _hwnd_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    _u32.GetClassNameW(hwnd, buf, 256)
    return buf.value

def _hwnd_title(hwnd):
    buf = ctypes.create_unicode_buffer(512)
    _u32.GetWindowTextW(hwnd, buf, 512)
    return buf.value

def _hwnd_child_texts(hwnd):
    texts = []
    buf   = ctypes.create_unicode_buffer(512)
    PROC  = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_size_t, ctypes.c_size_t)
    def _cb(child, _):
        _u32.GetWindowTextW(child, buf, 512)
        t = buf.value.strip()
        if t:
            texts.append(t)
        return True
    _u32.EnumChildWindows(hwnd, PROC(_cb), 0)
    return texts


def find_new_error_dialog(before_handles):
    """
    Proverava foreground prozor: ako je novi #32770 dialog (nije bio u snimku),
    cita mu naslov i tekst children-a i odlucuje da li je error za snimanje.

    Detektuje dva slucaja:
      1. 'File not found' — Outlook bug, zatvara sa Enter (OK)
      2. 'Confirm Save As' — overwrite dialog, zatvara sa Escape (No/odustani)

    Vraca True ako je dialog zatvoren i treba ponavljati, inace False.
    """
    FILE_NOT_FOUND_TITLE   = ["warning: this site can see edits you make"]
    FILE_NOT_FOUND_CONTENT = ["file not found", "check the file name"]
    OVERWRITE_TITLE        = ["confirm save as"]

    try:
        hwnd = _u32.GetForegroundWindow()
        if not hwnd:
            return False
        if _hwnd_class(hwnd) != "#32770":
            return False
        if hwnd in before_handles:
            return False  # postojeci prozor, nije novi

        title       = _hwnd_title(hwnd)
        child_texts = _hwnd_child_texts(hwnd)
        all_text    = (title + " " + " ".join(child_texts)).lower()
        print(f"  [DIAG] Foreground dialog: naslov='{title}' | sadrzaj={child_texts}")

        # Slucaj 1: File not found
        if any(m in title.lower() for m in FILE_NOT_FOUND_TITLE) or \
           any(m in all_text      for m in FILE_NOT_FOUND_CONTENT):
            print(f"  [SAVE] 'File not found' — zatvaramo (Enter)...")
            pyautogui.press("enter"); time.sleep(0.3)
            return True

        # Slucaj 2: Confirm Save As (overwrite)
        if any(m in title.lower() for m in OVERWRITE_TITLE):
            print(f"  [SAVE] 'Confirm Save As' (overwrite) — odbijamo (Escape)...")
            pyautogui.press("escape"); time.sleep(0.3)
            return True

        print(f"  [DIAG] Foreground dialog nije error za snimanje — ignorisemo.")
    except Exception as e:
        print(f"  [DIAG] Greska pri citanju foreground dialoga: {e}")
    return False


def find_save_dialog(timeout=120):
    """
    Ceka Save As dialog do `timeout` sekundi.
    Ako detektuje download prozor (Outlook preuzima mejl sa servera),
    stampa poruku i nastavlja da ceka — dialog ce se pojaviti posle downloada.
    """
    deadline        = time.time() + timeout
    download_noted  = False

    while time.time() < deadline:
        try:
            all_wins = Desktop(backend="win32").windows()
        except Exception:
            time.sleep(0.5)
            continue

        for w in all_wins:
            try:
                if w.class_name() != "#32770":
                    continue
                names = {(c.element_info.name or "").strip() for c in w.children()}
                if "&Save" in names or "Email Message (*.eml)" in names:
                    return w, "save"
                if any(n in names for n in ["&Yes", "Yes", "&No", "No", "Replace", "&Replace"]):
                    return w, "overwrite"
            except Exception:
                pass

        # Provjeri da li Outlook prikazuje download/progress prozor
        if not download_noted:
            for w in all_wins:
                try:
                    title = (w.window_text() or "").lower()
                    if any(k in title for k in ["download", "preuzimanje", "progress", "loading"]):
                        print(f"  [DOWNLOAD] Outlook preuzima mejl sa servera — cekam...")
                        toast("Preuzimanje mejla sa servera...\nCekam Save dialog.",
                              boja="#1a5276", trajanje=10)
                        download_noted = True
                        break
                except Exception:
                    pass

        time.sleep(0.5)
    return None, None


def get_subject_from_dialog(dialog):
    subject = "email"
    try:
        for child in dialog.children():
            name = (child.element_info.name or "").strip()
            if name.lower().endswith(".eml"):
                subject = name[:-4].strip()
                break
            if name.lower().endswith(" .eml"):
                subject = name[:-5].strip()
                break
            if name.lower().endswith(".msg"):
                # Teams pozivnica — Save As dialog nudi .msg, ali mi snimamo kao .eml
                subject = name[:-4].strip()
                break
    except Exception:
        pass
    subject = re.sub(r"^\d{4}-\d{2}-\d{2}_", "", subject).strip()
    return subject


def handle_save_dialog(date_str):
    dialog, dtype = find_save_dialog(timeout=6)

    if dialog is None:
        return None, "no_dialog", ""

    if dtype == "overwrite":
        # Overwrite dialog je vec otvoren pre navigacije — odbijamo i ponavljamo
        print(f"  [SAVE] Overwrite dialog pre navigacije — odbijamo (Escape)...")
        pyautogui.press("escape"); time.sleep(0.4)
        return None, "retry", ""

    subject  = get_subject_from_dialog(dialog)
    filename = sanitize_filename(f"{date_str}_{subject}") + ".eml"
    print(f"  Predmet : '{subject}'")
    print(f"  Fajl    : '{filename}'")

    # Upisujemo punu putanju (folder + naziv) direktno u polje za naziv fajla.
    # Windows Save As dijalog prihvata pun path i sam navigira u folder —
    # pouzdanije od Alt+D → folder → Alt+N → naziv (dva koraka).
    full_path = os.path.join(SAVE_FOLDER, filename)

    print(f"  Upisujem punu putanju u filename polje...")
    pyautogui.hotkey("alt", "n")
    time.sleep(_D_ALTN)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.10)
    pyperclip.copy(full_path)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.30)

    # Snapshot pre Enter — posle cemo traziti NOVE prozore
    dlg_snapshot = snapshot_dialogs()
    pyautogui.press("enter")

    # Cekaj da fajl fizicki bude na disku — proveravaj svakih 3s, timeout 60s
    filepath = os.path.join(SAVE_FOLDER, filename)
    print(f"  Cekam da se fajl pojavi na disku: {filename}")
    deadline = time.time() + 60
    saved = False
    while time.time() < deadline:
        if os.path.exists(filepath):
            saved = True
            print(f"  Fajl pronadjen na disku.")
            break
        # Proveri da li se pojavio novi error dialog (posle snimka)
        if find_new_error_dialog(dlg_snapshot):
            print(f"  [SAVE] Error dialog zatvoren — zatvaramo Save As i ponavljamo...")
            pyautogui.press("escape"); time.sleep(0.3)   # zatvori Save As
            return None, "retry", subject
        print(f"  ... jos nije snimljen, cekam 3s")
        time.sleep(3)

    if not saved:
        return None, "error:fajl nije snimljen u roku od 60 sekundi", subject

    return filename, "ok", subject


# ──────────────────────────────────────────────
# Move to Inbox2  (samo navigacija menija, bez fokus trikova)
# ──────────────────────────────────────────────

def move_to_inbox2(win):
    print("  [MOVE] Ctrl+Shift+V → Inbox2 → Enter")
    pyautogui.hotkey("ctrl", "shift", "v"); time.sleep(1.5)
    pyperclip.copy(DEST_FOLDER)
    pyautogui.hotkey("ctrl", "v"); time.sleep(0.5)
    pyautogui.press("enter"); time.sleep(1.0)
    print("  [MOVE] Premesten u Inbox2")


# ──────────────────────────────────────────────
# Hotkey handler
# ──────────────────────────────────────────────

def on_save_hotkey():
    global _pending_move, _batch_remaining, _batch_done, _pending_subject, _pending_file
    try:
        print(f"\n{'─'*50}")
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Hotkey")

        win = get_outlook_win_retry()
        if win is None:
            log_entry("", "", "error", "outlook not found posle 5 pokusaja")
            toast("Outlook nije pronadjen.\nBatch prekinut.", boja="#922b21", trajanje=5)
            return

        # ── MOD: premesti (automatski poziv posle snimanja) ──
        with _pending_move_lk:
            is_move = _pending_move
            _pending_move = False
            subj = _pending_subject
            fname = _pending_file

        if is_move:
            print(f"  [MODE] Premestanje u {DEST_FOLDER}...")
            navigate_to_top(win)
            move_to_inbox2(win)

            with _pending_move_lk:
                _batch_done += 1
                _batch_remaining -= 1
                done      = _batch_done
                remaining = _batch_remaining

            log_entry(subj, fname, "moved", DEST_FOLDER)
            print(f"  [BATCH] Obradjeno: {done}  Preostalo: {remaining}")

            if _stop_requested:
                with _pending_move_lk:
                    _batch_remaining = 0
                toast(f"Zaustavljeno na zahtev.\nObradjeno {done} mejlova.",
                      boja="#7d6608", trajanje=5)
                log_entry("", "", "batch_stopped", f"obradjeno {done}")
                print(f"  [STOP] Zaustavljeno — obradjeno {done} mejlova.")
            elif remaining > 0:
                toast(f"#{done} snimljeno i premesteno\nNastavljam...", boja="#1e8449", trajanje=2)
                time.sleep(0.5)
                threading.Thread(target=on_save_hotkey, daemon=True).start()
            else:
                toast(f"Batch gotov!\nObradjeno {done} mejlova.", boja="#1e8449", trajanje=5)
                log_entry("", "", "batch_done", f"obradjeno {done}")
                print(f"  [BATCH] Gotovo — {done} mejlova.")
            return

        # ── MOD: snimi ──
        date_str = datetime.now().strftime("%Y-%m-%d_%H%M%S")

        # Inicijalizuj batch brojace na prvom pozivu
        with _pending_move_lk:
            _batch_remaining = MAX_BATCH
            _batch_done      = 0
        print(f"  [BATCH] Start — do {MAX_BATCH} mejlova")

        GROUP_RETRY_WAIT  = 5
        filename     = None
        status       = None
        subject      = ""
        group_retries = 0

        for dialog_attempt in range(1, MAX_DIALOG_RETRY + MAX_GROUP_RETRY + 1):

            # Snapshot Downloads na pocetku svakog pokusaja — pre bilo kakve akcije
            dl_snapshot = snapshot_downloads()

            navigate_to_top(win)

            # ── Detekcija tipa: meeting/Teams pozivnica ili obican mejl ──
            is_meeting, item_name = is_meeting_item(win)
            print(f"  [TIP] is_meeting={is_meeting}  item={item_name[:60]!r}")

            if is_meeting:
                # ── MEETING: Shift+F10 → Save as EML ──
                if DRY_RUN:
                    toast("[TEST] Teams pozivnica — Shift+F10 tok", boja="#1e8449", trajanje=3)
                    return
                toast("Teams pozivnica — snimam kao EML...", boja="#1a5276", trajanje=8)
                open_saveas_via_contextmenu(win)
                filename, status, subject = handle_save_dialog(date_str)

            else:
                # ── OBICAN MEJL: Alt → F → A → Enter ──
                print("  [UIA] Alt → F...")
                pyautogui.press("alt"); time.sleep(0.10)
                pyautogui.press("f");   time.sleep(_D_FILE)

                # Pretrazi SVE prozore desktopu za "About Outlook"
                save_as_elem = None
                try:
                    for w in Desktop(backend="uia").windows():
                        try:
                            for elem in w.descendants():
                                try:
                                    name = (elem.element_info.name or "").strip()
                                    if name == "About Outlook":
                                        save_as_elem = elem
                                        print(f"  [UIA] Nasao: {name!r}")
                                        break
                                except Exception:
                                    pass
                            if save_as_elem:
                                break
                        except Exception:
                            pass
                except Exception as e:
                    print(f"  [UIA] Greska: {e}")

                if DRY_RUN:
                    pyautogui.press("escape"); time.sleep(0.4)
                    if save_as_elem:
                        toast("[TEST] File meni — About Outlook pronadjen", boja="#1e8449", trajanje=3)
                    else:
                        toast("[TEST] File meni nije otvoren — preskoceno", boja="#b05000", trajanje=3)
                    return

                if save_as_elem is None:
                    # File meni se nije otvorio — nije mejl, zatvori i pokusaj ponovo
                    pyautogui.press("escape"); time.sleep(0.4)
                    for _ in range(3):
                        pyautogui.press("escape"); time.sleep(1.0)
                    group_retries += 1
                    with _pending_move_lk:
                        done = _batch_done
                    if group_retries <= MAX_GROUP_RETRY:
                        print(f"  [UIA] Nije email meni (pokusaj {group_retries}/{MAX_GROUP_RETRY})"
                              f" — cekam 5s...")
                        toast(f"Nije mejl ({group_retries}/{MAX_GROUP_RETRY})\n"
                              f"Ponavljam za 5s...",
                              boja="#b05000", trajanje=5)
                        time.sleep(5)
                        continue
                    # Iscrpljeni svi pokusaji
                    with _pending_move_lk:
                        _batch_remaining = 0
                    if done > 0:
                        print(f"  [BATCH] Nema vise mejlova. Obradjeno: {done}")
                        toast(f"Nema vise mejlova.\nObradjeno: {done}", boja="#7d6608", trajanje=5)
                        log_entry("", "", "batch_done", f"obradjeno {done}")
                    else:
                        print("  !! Nije email meni — odustajem.")
                        toast("Nije mejl — klikni na mejl pa pokusaj ponovo.", boja="#b05000")
                        log_entry("", "", "skip", "nema File menija")
                    return

                # File meni otvoren — pritisni A (Save As) pa Enter
                toast("Snimam mejl...", boja="#1a5276", trajanje=8)
                print("  [SAVE] A → Enter (Save As)...")
                pyautogui.press("a");     time.sleep(0.10)
                pyautogui.press("enter"); time.sleep(_D_SA)

                filename, status, subject = handle_save_dialog(date_str)

            if status == "ok":
                break  # uspesno, izlazi iz retry petlje

            if status == "retry":
                # Outlook 'File not found' — dijalozi su vec zatvoreni, cekamo i ponavljamo
                print(f"  [RETRY] Outlook greska — cekam 10s pa ponavljam (pokusaj {dialog_attempt})...")
                toast("Outlook greska pri snimanju.\nPonovni pokusaj za 10s...",
                      boja="#b05000", trajanje=10)
                time.sleep(10)
                continue

            if status == "no_dialog":
                # Nema Save dialoga — probaj da uzmemo fajl iz Downloads
                print("  [DOWNLOAD] Nema Save dialoga — trazim fajl u Downloads...")
                toast("Nema Save dialoga.\nTrazim fajl u Downloads...",
                      boja="#1a5276", trajanje=5)
                filename, status, subject = handle_download_file(dl_snapshot, date_str)
                if status == "ok":
                    break  # uspesno — nastavi normalno

                # Ni Downloads nije pomogao
                log_entry("", "", "no_dialog", f"pokusaj {dialog_attempt}/{MAX_DIALOG_RETRY}")
                if dialog_attempt < MAX_DIALOG_RETRY:
                    print(f"  !! Nema fajla (pokusaj {dialog_attempt}/{MAX_DIALOG_RETRY})"
                          f" — cekam {DIALOG_RETRY_WAIT}s...")
                    toast(f"Nema fajla ({dialog_attempt}/{MAX_DIALOG_RETRY})\n"
                          f"Ponovni pokusaj za {DIALOG_RETRY_WAIT}s...",
                          boja="#b05000", trajanje=DIALOG_RETRY_WAIT)
                    time.sleep(DIALOG_RETRY_WAIT)
                    continue
                else:
                    print(f"  !! Nema fajla ni posle {MAX_DIALOG_RETRY} pokusaja — odustajem.")
                    toast(f"Nema fajla.\nBatch prekinut.", boja="#922b21", trajanje=5)
                    with _pending_move_lk:
                        _batch_remaining = 0
                    return
            else:
                # Druga greska (nije no_dialog) — ne ponavljaj
                with _pending_move_lk:
                    _batch_remaining = 0
                print(f"  !! Greska: {status}")
                log_entry(subject, "", "error", status)
                toast(f"Greska:\n{status}", boja="#922b21")
                return

        if status == "ok":
            print(f"  OK Snimljeno: {filename}")
            log_entry(subject, filename, "ok")

            # Postavi flag i automatski pokreni "move" fazu
            with _pending_move_lk:
                _pending_move    = True
                _pending_subject = subject
                _pending_file    = filename
            time.sleep(0.5)
            print("  Pokrecim move fazu...")
            threading.Thread(target=on_save_hotkey, daemon=True).start()

    except Exception as e:
        print(f"  !! Neocekivana greska: {e}")
        log_entry("", "", "error", str(e))


# ──────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────

def main():
    global _log_file, _log_writer, DEST_FOLDER

    print("=" * 55)
    print("Outlook EML Saver — hotkey mod")
    print(f"  Folder : {SAVE_FOLDER}")
    print(f"  Log    : {LOG_FILE}")
    print(f"  Snimi  : {HOTKEY_SAVE.upper()}")
    print("=" * 55)
    print()

    # Izbor odredisnog foldera
    opcije = [x.strip() for x in _cfg.get("behaviour", "dest_folder_options").split(",")]
    print("  Odredisni folder za premestanje:")
    for i, naziv in enumerate(opcije, 1):
        default_tag = " (default)" if naziv == DEST_FOLDER else ""
        print(f"    {i}. {naziv}{default_tag}")
    print(f"    5. Unesi drugi naziv")
    print()
    izbor = input("  Izbor [1]: ").strip()

    if izbor == "" or izbor == "1":
        DEST_FOLDER = opcije[0]
    elif izbor == "2":
        DEST_FOLDER = opcije[1]
    elif izbor == "3":
        DEST_FOLDER = opcije[2]
    elif izbor == "4":
        DEST_FOLDER = opcije[3]
    elif izbor == "5":
        unos = input("  Naziv foldera: ").strip()
        if unos:
            DEST_FOLDER = unos
    else:
        # ako je upisao direktno naziv
        DEST_FOLDER = izbor

    print(f"  Premesta u: {DEST_FOLDER}")
    print()

    if not os.path.isdir(SAVE_FOLDER):
        sys.exit(f"GRESKA: Folder ne postoji: {SAVE_FOLDER}")

    _log_file, _log_writer = init_log()

    def on_hotkey():
        global _stop_requested
        with _pending_move_lk:
            batch_aktivan = _batch_remaining > 0 or _pending_move
        if batch_aktivan:
            _stop_requested = True
            print("\n  [STOP] Zahtev za zaustavljanje — zavrsavam tekuci mejl...")
            toast("Zaustavljam posle tekuceg mejla...", boja="#7d6608", trajanje=3)
        else:
            _stop_requested = False
            threading.Thread(target=on_save_hotkey, daemon=True).start()

    keyboard.add_hotkey(HOTKEY_SAVE, on_hotkey, suppress=True)

    # Na startu: prebaci fokus na gornji mejl
    win = get_outlook_win_retry()
    if win:
        print("  Navigiram na gornji mejl...")
        navigate_to_top(win)
    else:
        print("  Outlook nije pronadjen pri startu.")

    if DRY_RUN:
        print(f"  *** DRY RUN — nista se ne snima ***")
    print(f"Slusa... (Enter = kraj)")

    try:
        input()
    except KeyboardInterrupt:
        pass
    finally:
        keyboard.unhook_all()
        _log_file.close()
        print("Zavrseno.")
        _tee_out.close_log()
        _tee_err.close_log()


if __name__ == "__main__":
    main()
