# app.py â€“ DriverBell v1.7.2 (UX Enhanced, Tray & Autostart Fixed)
# Developed by Dwi Citolaksono â€” 8 November 2025
import os, sys, csv, json, uuid, time, datetime as dt
from dataclasses import dataclass, asdict
from pathlib import Path
import threading
import ctypes
from urllib.parse import unquote, urlparse
import wx, wx.adv
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
import pytz
from tzlocal import get_localzone

_AUDIO_BACKEND_INITIALIZED = False
_AUDIO_IMPORT_LOCK = threading.Lock()
_AUDIO_MIXER_LOCK = threading.Lock()

# BASS backend globals
_BASS = None
_BASS_DEVICE_INDEX = -999
_BASS_STREAM = 0
_BASS_INIT_ERROR = ""

BASS_UNICODE = 0x80000000
BASS_STREAM_PRESCAN = 0x20000

# --- optional pywin / COM for autostart shortcut
try:
    import pythoncom
    from win32com.shell import shell, shellcon
    has_pywin = True
except Exception:
    has_pywin = False

APP_NAME = "DriverBell"
VERSION = "1.7.2"
RELEASE_DATE = "8 November 2025"
DEFAULT_TZ = "Asia/Jakarta"

# === PATCH: user-writable paths (ganti CONFIG_FILE/LOG_FILE lama) ===
APPDATA_DIR = Path(os.getenv("APPDATA") or (Path.home() / "AppData" / "Roaming")) / "DriverBell"
APPDATA_DIR.mkdir(parents=True, exist_ok=True)

CONFIG_PATH = APPDATA_DIR / "bell_config.json"
LOG_PATH    = APPDATA_DIR / "bell_log.csv"

def _migrate_legacy_files():
    """Pindahkan file lama dari folder kerja ke %APPDATA% jika ada."""
    legacy_cfg = Path("bell_config.json")
    legacy_log = Path("bell_log.csv")
    try:
        if legacy_cfg.exists() and not CONFIG_PATH.exists():
            CONFIG_PATH.write_text(legacy_cfg.read_text(encoding="utf-8"), encoding="utf-8")
        if legacy_log.exists() and not LOG_PATH.exists():
            LOG_PATH.write_text(legacy_log.read_text(encoding="utf-8", errors="ignore"), encoding="utf-8")
    except Exception:
        pass
# === END PATCH ===

LANG_STRINGS = {
    "id": {
        "add":"Tambahkan Bell (&T)","settings":"Pengaturan (&S)","log":"Lihat Log (&L)","exit":"Keluar (&X)",
        "activate":"Aktifkan","deactivate":"Nonaktifkan","edit":"Edit","delete":"Hapus",
        "ready":"Siap. Tab untuk navigasi. Alt+T tambah, Alt+S pengaturan, Alt+L log, Alt+A tentang, Alt+X keluar.",
        "timezone":"Zona Waktu","language":"Bahasa","run_in_background":"Jalankan bell saat jendela ditutup (tray)",
        "output_device":"Perangkat Keluaran","output_device_default":"Default Sistem",
        "autostart":"Jalankan otomatis saat Windows menyala","show_welcome_on_startup":"Tampilkan dialog sambutan saat aplikasi dibuka","save":"Simpan","cancel":"Batal","show":"Tampilkan","quit":"Keluar",
        "tray_title":"Driver Bell (Berjalan)","about":"Tentang","confirm_delete":"Hapus '{name}' ?","no_audio":"File audio belum dipilih atau tidak ditemukan.",
        "added":"Ditambahkan:","updated":"Diperbarui","deleted":"Dihapus","activated":"Diaktifkan","deactivated":"Dinonaktifkan","settings_saved":"Pengaturan disimpan",
        "menu_help":"Bantuan","menu_about":"Tentang DriverBell",
        "about_text":(
            "Selamat datang di DriverBell versi 1.7.2!\n\n"
            "Software ini adalah software pemutaran bell sekolah otomatis yang dibuat oleh "
            "Dwi Citolaksono (Seorang Tunanetra), sebagai wujud bakti dan inovasi dalam dunia pendidikan.\n\n"
            "Software ini dikembangkan dalam bahasa pemrograman Python, didedikasikan bagi seluruh sekolah "
            "yang membutuhkan software ini, tak terbatas siapapun dan di mana pun.\n\n"
            "Berbeda dengan software pemutar bell otomatis lainnya, software ini dirancang agar dapat berjalan "
            "secara offline, mengedepankan aksesibilitas dan stabilitas, serta dapat berjalan walau dengan "
            "komputer yang memiliki spesifikasi minimum.\n\n"
            "Mari turut berdonasi, untuk membantu keberlanjutan pengembangan software ini."
        ),
        "entry_name":"Nama Entry:","start_time":"Jam Mulai (HH:MM):","end_time":"Jam Akhir (opsional):","audio_file":"File Audio Bell:",
        "browse":"Jelajahi berkas audio","preview_audio":"Pratinjau audio bell","stop_preview":"Hentikan pratinjau bell","audio_pick":"Gagal memutar pratinjau audio.","wav_only":"File audio harus berformat WAV (.wav).","days_label":"Hari:","name_required":"Nama tidak boleh kosong.","select_day":"Pilih minimal satu hari.",
        "col_name":"Nama","col_start":"Mulai","col_end":"Akhir","col_days":"Hari","col_active":"Status","col_last_played":"Terakhir Diputar","updated_label":"Segarkan",
        "status_active":"Aktif","status_inactive":"Nonaktif","last_played_empty":"Belum pernah diputar",
        "days_list":["Senin","Selasa","Rabu","Kamis","Jumat","Sabtu","Minggu"]
    },
    "en": {
        "add":"Add Bell (&T)","settings":"Settings (&S)","log":"View Log (&L)","exit":"Exit (&X)",
        "activate":"Activate","deactivate":"Deactivate","edit":"Edit","delete":"Delete",
        "ready":"Ready. Tab to navigate. Alt+T add, Alt+S settings, Alt+L log, Alt+A about, Alt+X exit.",
        "timezone":"Timezone","language":"Language","run_in_background":"Keep bells running when window closed (tray)",
        "output_device":"Output Device","output_device_default":"System Default",
        "autostart":"Auto-start when Windows starts","show_welcome_on_startup":"Show welcome dialog on startup","save":"Save","cancel":"Cancel","show":"Show","quit":"Quit",
        "tray_title":"Driver Bell (Running)","about":"About","confirm_delete":"Delete '{name}' ?","no_audio":"Audio file not selected or missing.",
        "added":"Added:","updated":"Updated","deleted":"Deleted","activated":"Activated","deactivated":"Deactivated","settings_saved":"Settings saved",
        "menu_help":"Help","menu_about":"About DriverBell",
        "about_text":(
            "Welcome to DriverBell version 1.7.2!\n\n"
            "This software is an automatic school bell playback system created by "
            "Dwi Citolaksono (a blind developer), as a form of dedication and innovation in education.\n\n"
            "This software is developed in Python and dedicated to all schools that need it, "
            "without limitation of who they are or where they are.\n\n"
            "Unlike many other automatic bell players, this software is designed to run offline, "
            "prioritizing accessibility and stability, and it can run even on computers with minimum specifications.\n\n"
            "Please consider donating to support the ongoing development of this software."
        ),
        "entry_name":"Entry Name:","start_time":"Start Time (HH:MM):","end_time":"End Time (optional):","audio_file":"Audio File:",
        "browse":"Browse Audio File","preview_audio":"Preview the bell","stop_preview":"Stop the bell preview","audio_pick":"Failed to play audio preview.","wav_only":"Audio file must be in WAV format (.wav).","days_label":"Days:","name_required":"Name cannot be empty.","select_day":"Select at least one day.",
        "col_name":"Name","col_start":"Start","col_end":"End","col_days":"Days","col_active":"Status","col_last_played":"Last Played","updated_label":"Refresh",
        "status_active":"Active","status_inactive":"Inactive","last_played_empty":"Never played",
        "days_list":["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    }
}
WELCOME_TEXT = {
    "id": {
        "title": "Welcome to DriverBell versi 1.7.2!",
        "message": (
            "DriverBell dikembangkan oleh Dwi Citolaksono untuk mendukung sekolah "
            "dan lembaga pendidikan. Jika Anda ingin membantu keberlangsungan "
            "pengembangan software ini, donasi Anda sangat kami apresiasi."
        ),
        "ok": "OK",
        "donate": "Donasi sekarang",
        "dont_show": "Jangan tampilkan pesan ini lagi",
    },
    "en": {
        "title": "Welcome to DriverBell Version 1.7.2!",
        "message": (
            "DriverBell is developed by Dwi Citolaksono to support schools and "
            "educational institutions. If you would like to help sustain the "
            "development of this software, your donation is greatly appreciated."
        ),
        "ok": "OK",
        "donate": "Donate now",
        "dont_show": "Don't show this message again",
    },
}

@dataclass
class BellEntry:
    id:str; name:str; start_time:str; end_time:str; audio_path:str; days:list; active:bool; last_played:str=""

@dataclass
class AppConfig:
    timezone: str = DEFAULT_TZ
    language: str = "en"
    run_in_background: bool = False
    autostart: bool = False
    show_welcome: bool = True
    output_device: str = ""
    entries: list = None

    def to_json(self):
        return {
            "timezone": self.timezone,
            "language": self.language,
            "run_in_background": self.run_in_background,
            "autostart": self.autostart,
            "show_welcome": self.show_welcome,
            "output_device": self.output_device,
            "entries": [asdict(e) for e in (self.entries or [])],
        }

    @staticmethod
    def from_json(data):
        cfg = AppConfig()
        cfg.timezone = data.get("timezone", DEFAULT_TZ)
        cfg.language = data.get("language", "en")
        cfg.run_in_background = bool(data.get("run_in_background", False))
        cfg.autostart = bool(data.get("autostart", False))
        cfg.show_welcome = bool(data.get("show_welcome", True))
        cfg.output_device = str(data.get("output_device", "") or "")
        cfg.entries = []
        for e in data.get("entries", []):
            cfg.entries.append(BellEntry(
                id=e.get("id", str(uuid.uuid4())),
                name=e.get("name", "Untitled"),
                start_time=e.get("start_time", "07:00"),
                end_time=e.get("end_time", ""),
                audio_path=e.get("audio_path", ""),
                days=e.get("days", [0, 1, 2, 3, 4]),
                active=bool(e.get("active", False)),
                last_played=e.get("last_played", "")
            ))
        return cfg

def load_config():
    try:
        if CONFIG_PATH.exists():
            return AppConfig.from_json(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
    except Exception:
        pass
    return AppConfig(timezone=DEFAULT_TZ, entries=[])

def save_config(cfg: AppConfig):
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(cfg.to_json(), ensure_ascii=False, indent=2), encoding="utf-8")

def append_log(row: dict):
    header = ["timestamp", "name", "action", "result", "details"]
    exists = LOG_PATH.exists()
    with LOG_PATH.open("a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=header)
        if not exists:
            w.writeheader()
        w.writerow(row)

def _normalize_audio_path(path: str) -> str:
    if not path:
        return ""
    value = str(path).strip().strip('"').strip("'")
    if value.lower().startswith("file://"):
        try:
            parsed = urlparse(value)
            value = unquote(parsed.path or "")
            if os.name == "nt" and value.startswith("/") and len(value) > 2 and value[2] == ":":
                value = value[1:]
        except Exception:
            pass
    value = os.path.expandvars(os.path.expanduser(value))
    return os.path.normpath(value)

def _resolve_existing_audio_path(path: str) -> str:
    base = _normalize_audio_path(path)
    if not base:
        return ""
    # Coba beberapa variasi yang sering muncul dari input UI/user.
    candidates = [base]
    abs_base = os.path.abspath(base)
    if abs_base not in candidates:
        candidates.append(abs_base)
    for c in list(candidates):
        c2 = c.rstrip(" .")
        if c2 and c2 not in candidates:
            candidates.append(c2)
    for c in candidates:
        try:
            if os.path.isfile(c):
                return c
            with open(c, "rb"):
                return c
        except Exception:
            continue
    # Fallback: cari nama file secara case-insensitive di folder parent.
    try:
        parent = Path(base).parent
        name = Path(base).name.lower()
        if parent.exists():
            for f in parent.iterdir():
                if f.is_file() and f.name.lower() == name:
                    return str(f)
    except Exception:
        pass
    return ""

class BASS_DEVICEINFO(ctypes.Structure):
    _fields_ = [
        ("name", ctypes.c_char_p),
        ("driver", ctypes.c_char_p),
        ("flags", ctypes.c_uint),
    ]

def _decode_bass_cstr(value) -> str:
    if not value:
        return ""
    try:
        return value.decode("utf-8", "ignore").strip()
    except Exception:
        try:
            return value.decode("mbcs", "ignore").strip()
        except Exception:
            return str(value).strip()

def _bass_dll_candidates():
    candidates = []
    if hasattr(sys, "_MEIPASS"):
        candidates.append(Path(sys._MEIPASS) / "bass.dll")
        candidates.append(Path(sys._MEIPASS) / "assets" / "audio" / "bass.dll")
    exe_dir = Path(sys.executable).resolve().parent
    candidates.append(exe_dir / "bass.dll")
    candidates.append(exe_dir / "_internal" / "bass.dll")
    candidates.append(Path(__file__).resolve().parent / "assets" / "audio" / "bass.dll")
    return candidates

def _load_bass():
    global _BASS, _BASS_INIT_ERROR
    if _BASS is not None:
        return True
    with _AUDIO_IMPORT_LOCK:
        if _BASS is not None:
            return True
        for cand in _bass_dll_candidates():
            try:
                if cand.exists():
                    _BASS = ctypes.WinDLL(str(cand))
                    break
            except Exception:
                continue
        if _BASS is None:
            _BASS_INIT_ERROR = "bass.dll not found"
            return False
        try:
            _BASS.BASS_Init.argtypes = [ctypes.c_int, ctypes.c_uint, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p]
            _BASS.BASS_Init.restype = ctypes.c_int
            _BASS.BASS_Free.argtypes = []
            _BASS.BASS_Free.restype = ctypes.c_int
            _BASS.BASS_GetDeviceInfo.argtypes = [ctypes.c_uint, ctypes.POINTER(BASS_DEVICEINFO)]
            _BASS.BASS_GetDeviceInfo.restype = ctypes.c_int
            # NOTE: argtypes sengaja tidak dipatok untuk BASS_StreamCreateFile agar
            # path Unicode/ANSI bisa dipass aman via ctypes pada runtime Windows.
            _BASS.BASS_StreamCreateFile.restype = ctypes.c_uint
            _BASS.BASS_ChannelPlay.argtypes = [ctypes.c_uint, ctypes.c_int]
            _BASS.BASS_ChannelPlay.restype = ctypes.c_int
            _BASS.BASS_ChannelIsActive.argtypes = [ctypes.c_uint]
            _BASS.BASS_ChannelIsActive.restype = ctypes.c_int
            _BASS.BASS_ChannelStop.argtypes = [ctypes.c_uint]
            _BASS.BASS_ChannelStop.restype = ctypes.c_int
            _BASS.BASS_StreamFree.argtypes = [ctypes.c_uint]
            _BASS.BASS_StreamFree.restype = ctypes.c_int
            _BASS.BASS_ErrorGetCode.argtypes = []
            _BASS.BASS_ErrorGetCode.restype = ctypes.c_int
            return True
        except Exception as ex:
            _BASS_INIT_ERROR = f"BASS bind error: {ex}"
            _BASS = None
            return False

def _bass_error_text(prefix="BASS"):
    if _BASS is None:
        return _BASS_INIT_ERROR or f"{prefix} unavailable"
    try:
        return f"{prefix} error {_BASS.BASS_ErrorGetCode()}"
    except Exception:
        return f"{prefix} error unknown"

def _list_bass_devices():
    items = []
    if not _load_bass():
        return items
    i = 0
    while True:
        info = BASS_DEVICEINFO()
        ok = _BASS.BASS_GetDeviceInfo(i, ctypes.byref(info))
        if not ok:
            break
        name = _decode_bass_cstr(info.name)
        if name:
            items.append((i, name))
        i += 1
    return items

def _resolve_bass_device_index(selected_name: str):
    selected = (selected_name or "").strip()
    if not selected:
        return -1
    devices = _list_bass_devices()
    if not devices:
        return -1
    for idx, name in devices:
        if name == selected:
            return idx
    s = selected.lower()
    for idx, name in devices:
        n = name.lower()
        if s in n or n in s:
            return idx
    return -1

def initialize_audio_backend():
    global _AUDIO_BACKEND_INITIALIZED, _BASS_INIT_ERROR
    if _AUDIO_BACKEND_INITIALIZED:
        return _BASS is not None
    ok = _load_bass()
    _AUDIO_BACKEND_INITIALIZED = True
    if not ok and not _BASS_INIT_ERROR:
        _BASS_INIT_ERROR = "BASS initialization failed"
    return ok

def list_audio_output_devices():
    return [name for _, name in _list_bass_devices()]

class AudioPlayer:
    """
    Pemutar audio multi-format berbasis BASS.
    Bunyi sekali (one-shot), non-blocking.
    """
    def __init__(self, output_device: str = ""):
        self.output_device = (output_device or "").strip()
        self.last_error = ""

    def _ensure_bass(self) -> bool:
        global _BASS_DEVICE_INDEX
        if not initialize_audio_backend():
            self.last_error = _BASS_INIT_ERROR or "BASS not available"
            return False
        try:
            with _AUDIO_MIXER_LOCK:
                target_idx = _resolve_bass_device_index(self.output_device) if self.output_device else -1
                if self.output_device and target_idx < 0:
                    self.last_error = f"Selected output device not available: {self.output_device}"
                    return False
                if _BASS_DEVICE_INDEX == target_idx:
                    return True
                try:
                    _BASS.BASS_Free()
                except Exception:
                    pass
                ok = _BASS.BASS_Init(target_idx, 44100, 0, None, None)
                if not ok:
                    self.last_error = _bass_error_text("BASS_Init")
                    return False
                _BASS_DEVICE_INDEX = target_idx
                return True
        except Exception as ex:
            self.last_error = f"BASS init failed: {ex}"
            return False

    def play_loop(self, path: str) -> bool:
        global _BASS_STREAM
        resolved_path = _resolve_existing_audio_path(path)
        self.last_error = ""
        if not resolved_path:
            self.last_error = "Audio file not found"
            return False
        if not self._ensure_bass():
            return False
        try:
            with _AUDIO_MIXER_LOCK:
                if _BASS_STREAM:
                    try:
                        _BASS.BASS_ChannelStop(_BASS_STREAM)
                        _BASS.BASS_StreamFree(_BASS_STREAM)
                    except Exception:
                        pass
                    _BASS_STREAM = 0
                stream = _BASS.BASS_StreamCreateFile(
                    0,
                    ctypes.c_wchar_p(resolved_path),
                    0,
                    0,
                    BASS_UNICODE | BASS_STREAM_PRESCAN,
                )
                if not stream:
                    self.last_error = _bass_error_text("BASS_StreamCreateFile")
                    return False
                ok = _BASS.BASS_ChannelPlay(stream, 0)
                if not ok:
                    _BASS.BASS_StreamFree(stream)
                    self.last_error = _bass_error_text("BASS_ChannelPlay")
                    return False
                _BASS_STREAM = stream
            return True
        except Exception as ex:
            self.last_error = f"BASS playback error: {ex}"
            return False

    def stop(self):
        global _BASS_STREAM
        try:
            with _AUDIO_MIXER_LOCK:
                if _BASS_STREAM:
                    _BASS.BASS_ChannelStop(_BASS_STREAM)
                    _BASS.BASS_StreamFree(_BASS_STREAM)
                    _BASS_STREAM = 0
        except Exception:
            pass

    def is_playing(self) -> bool:
        try:
            with _AUDIO_MIXER_LOCK:
                if not _BASS_STREAM or _BASS is None:
                    return False
                return bool(_BASS.BASS_ChannelIsActive(_BASS_STREAM) == 1)
        except Exception:
            return False

class BellScheduler:
    def __init__(self,tz_name:str,notify_cb,mark_last_played_cb=None,output_device_getter=None):
        self.tz=pytz.timezone(tz_name)
        self.sched=BackgroundScheduler(timezone=self.tz)
        self.sched.start()
        self.notify_cb=notify_cb
        self.mark_last_played_cb=mark_last_played_cb
        self.output_device_getter=output_device_getter
        self.players={}
    def shutdown(self):
        try:
            self.sched.shutdown(wait=False)
        except Exception:
            pass
    def clear_jobs_for(self,entry_id:str):
        for job in list(self.sched.get_jobs()):
            if str(job.id).startswith(entry_id+"::"):
                try:
                    self.sched.remove_job(job.id)
                except Exception:
                    pass
    def schedule_entry(self,e:BellEntry):
        try:
            hh,mm=map(int,e.start_time.split(":"))
        except Exception:
            return
        dow=",".join(str(d) for d in e.days)
        self.sched.add_job(self._start, CronTrigger(day_of_week=dow,hour=hh,minute=mm),
                           id=f"{e.id}::start", args=[e], replace_existing=True)
        if e.end_time:
            try:
                eh,em=map(int,e.end_time.split(":"))
                self.sched.add_job(self._stop, CronTrigger(day_of_week=dow,hour=eh,minute=em),
                                   id=f"{e.id}::stop", args=[e], replace_existing=True)
            except Exception:
                pass
    def _start(self,e:BellEntry):
        ok,details=False,""
        played_at=dt.datetime.now(self.tz)
        try:
            output_device = ""
            if callable(self.output_device_getter):
                output_device = self.output_device_getter() or ""
            p=AudioPlayer(output_device); self.players[e.id]=p
            ok=p.play_loop(e.audio_path)
            if not ok:
                details = p.last_error or "Audio file tidak tersedia atau tidak dapat diputar"
        except Exception as ex:
            ok,details=False,str(ex)
        append_log({
            "timestamp": played_at.isoformat(),
            "name": e.name,
            "action": "START",
            "result": "OK" if ok else "FAIL",
            "details": details
        })
        if ok and callable(self.mark_last_played_cb):
            try:
                self.mark_last_played_cb(e.id, played_at.isoformat())
            except Exception:
                pass
        self.notify_cb("PLAY_START", e.name, ok, details)
    def _stop(self,e:BellEntry):
        ok,details=True,""
        try:
            p=self.players.get(e.id)
            if p: p.stop()
        except Exception as ex:
            ok,details=False,str(ex)
        append_log({"timestamp":dt.datetime.now(self.tz).isoformat(),"name":e.name,"action":"STOP","result":"OK" if ok else "FAIL","details":details})
        self.notify_cb("PLAY_STOP",e.name,ok,details)

class MainFrame(wx.Frame):
    def __init__(self,cfg:AppConfig):
        super().__init__(None,title=APP_NAME,size=(960,560))
        self.cfg=cfg
        self._really_quit=False  # flag untuk keluar total
        self.strings=LANG_STRINGS.get(self.cfg.language,LANG_STRINGS["en"])
        self.scheduler=BellScheduler(
            self.cfg.timezone or DEFAULT_TZ,
            self._notify,
            self._mark_last_played,
            lambda: getattr(self.cfg, "output_device", "")
        )
        self._brand_icon=self._build_brand_icon((32,32))
        if self._brand_icon.IsOk():
            self.SetIcon(self._brand_icon)

        panel=wx.Panel(self)
        panel.SetBackgroundColour(wx.Colour(244, 238, 222))
        vbox=wx.BoxSizer(wx.VERTICAL)

        hs=wx.BoxSizer(wx.VERTICAL)
        self.header_title=wx.StaticText(panel,label="DriverBell")
        self.header_subtitle=wx.StaticText(panel,label=self._header_subtitle_text())
        self.header_title.SetForegroundColour(wx.Colour(35, 78, 46))
        self.header_subtitle.SetForegroundColour(wx.Colour(73, 91, 62))
        self.header_title.SetFont(wx.Font(16, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        hs.Add(self.header_title,0,wx.LEFT|wx.TOP,10)
        hs.Add(self.header_subtitle,0,wx.LEFT|wx.BOTTOM,10)
        vbox.Add(hs,0,wx.EXPAND|wx.LEFT|wx.RIGHT|wx.TOP,10)
        vbox.Add(wx.StaticLine(panel),0,wx.EXPAND|wx.LEFT|wx.RIGHT,10)

        hb=wx.BoxSizer(wx.HORIZONTAL)
        self.btn_add=wx.Button(panel,label=self.strings["add"])
        self.btn_set=wx.Button(panel,label=self.strings["settings"])
        self.btn_log=wx.Button(panel,label=self.strings["log"])
        self.btn_about=wx.Button(panel,label=self.strings.get("about","About"))
        self.btn_exit=wx.Button(panel,label=self.strings["exit"])
        for b in (self.btn_add,self.btn_set,self.btn_log,self.btn_about,self.btn_exit):
            hb.Add(b,0,wx.RIGHT,6)
        vbox.Add(hb,0,wx.ALL,10)

        tb=wx.BoxSizer(wx.HORIZONTAL)
        self.btn_activate=wx.Button(panel,label=self.strings["activate"])
        self.btn_deactivate=wx.Button(panel,label=self.strings["deactivate"])
        self.btn_edit=wx.Button(panel,label=self.strings["edit"])
        self.btn_delete=wx.Button(panel,label=self.strings["delete"])
        for b in (self.btn_activate,self.btn_deactivate,self.btn_edit,self.btn_delete):
            tb.Add(b,0,wx.RIGHT,6)
        vbox.Add(tb,0,wx.LEFT|wx.RIGHT|wx.BOTTOM,10)

        self.list=wx.ListCtrl(panel,style=wx.LC_REPORT|wx.BORDER_SUNKEN)
        self.list.SetBackgroundColour(wx.Colour(255, 252, 240))
        self.list.SetMinSize((-1, 220))
        cols=[self.strings["col_name"],self.strings["col_start"],self.strings["col_end"],self.strings["col_days"],self.strings["col_active"],self.strings["col_last_played"]]
        for i,c in enumerate(cols):
            self.list.InsertColumn(i,c)
        vbox.Add(self.list,1,wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM,10)

        self.statusbar=self.CreateStatusBar()
        self.SetStatusText(self.strings["ready"])
        panel.SetSizer(vbox)
        self.Centre()

        self.popup=wx.Menu()
        self.popup_act=self.popup.Append(501,self.strings["activate"])
        self.popup_edit=self.popup.Append(503,self.strings["edit"])
        self.popup_delete=self.popup.Append(504,self.strings["delete"])
        self.popup_log=self.popup.Append(505,self.strings["log"])

        # Global accelerator IDs
        self.ID_ADD=wx.NewIdRef(); self.ID_SETTINGS=wx.NewIdRef(); self.ID_LOG=wx.NewIdRef(); self.ID_ABOUT=wx.NewIdRef(); self.ID_EXIT=wx.NewIdRef()

        # bindings
        self.Bind(wx.EVT_BUTTON,lambda e: self._add(),self.btn_add)
        self.Bind(wx.EVT_BUTTON,lambda e: self._open_settings(),self.btn_set)
        self.Bind(wx.EVT_BUTTON,lambda e: self._open_log(),self.btn_log)
        self.Bind(wx.EVT_BUTTON,lambda e: self._show_about(),self.btn_about)

        self.Bind(wx.EVT_BUTTON,lambda e: self._edit(),self.btn_edit)
        self.Bind(wx.EVT_BUTTON,lambda e: self._delete(),self.btn_delete)
        self.Bind(wx.EVT_BUTTON,lambda e: self._activate(),self.btn_activate)
        self.Bind(wx.EVT_BUTTON,lambda e: self._deactivate(),self.btn_deactivate)

        # PATCH: Exit -> _quit() agar set flag dan memicu EVT_CLOSE
        self.Bind(wx.EVT_BUTTON,lambda e: self._quit(),self.btn_exit)

        for mid,fn in [(501,self._toggle_active),(503,self._edit),(504,self._delete),(505,self._open_log)]:
            self.Bind(wx.EVT_MENU, lambda e,f=fn: f(), id=mid)

        # bind global accelerator menu ids to functions
        self.Bind(wx.EVT_MENU, lambda e: self._add(), id=self.ID_ADD.GetId())
        self.Bind(wx.EVT_MENU, lambda e: self._open_settings(), id=self.ID_SETTINGS.GetId())
        self.Bind(wx.EVT_MENU, lambda e: self._open_log(), id=self.ID_LOG.GetId())
        self.Bind(wx.EVT_MENU, lambda e: self._show_about(), id=self.ID_ABOUT.GetId())
        self.Bind(wx.EVT_MENU, lambda e: self._quit(), id=self.ID_EXIT.GetId())
        self.SetAcceleratorTable(wx.AcceleratorTable([
            (wx.ACCEL_ALT, ord('T'), self.ID_ADD.GetId()),
            (wx.ACCEL_ALT, ord('S'), self.ID_SETTINGS.GetId()),
            (wx.ACCEL_ALT, ord('L'), self.ID_LOG.GetId()),
            (wx.ACCEL_ALT, ord('A'), self.ID_ABOUT.GetId()),
            (wx.ACCEL_ALT, ord('X'), self.ID_EXIT.GetId()),
        ]))

        self.list.Bind(wx.EVT_LIST_ITEM_ACTIVATED, lambda e: self._edit())
        self.list.Bind(wx.EVT_CONTEXT_MENU,self._on_show_popup)
        self.list.Bind(wx.EVT_LIST_ITEM_RIGHT_CLICK,self._on_show_popup)
        self.list.Bind(wx.EVT_KEY_DOWN,self._on_list_key)
        self.list.Bind(wx.EVT_CHAR_HOOK,self._on_list_char_hook)
        self.list.Bind(wx.EVT_SET_FOCUS,self._on_list_focus)
        self.list.Bind(wx.EVT_LIST_ITEM_SELECTED,self._on_list_selected)
        self.list.Bind(wx.EVT_LIST_ITEM_FOCUSED,self._on_list_focused)
        self.Bind(wx.EVT_CHAR_HOOK,self._on_frame_char_hook)

        # PATCH: tangani EVT_CLOSE agar hide to tray saat run_in_background aktif
        self.Bind(wx.EVT_CLOSE, self._on_close)

        self.tray_icon=None
        self._refresh_list()
        self._apply_school_theme()
        self._apply_button_icons()
        self._apply_tab_order()
        self.btn_add.SetFocus()
        self.Show()

        for e in (self.cfg.entries or []):
            if e.active:
                self.scheduler.schedule_entry(e)
        self._show_welcome_if_needed()

    def _quit(self):
        self._really_quit = True
        # memicu EVT_CLOSE; handler _on_close akan mengurus cleanup
        self.Close()

    def _show_welcome_if_needed(self):
        # Jika user sudah memilih untuk tidak menampilkan lagi, langsung keluar
        if not getattr(self.cfg, "show_welcome", True):
            return

        lang = self.cfg.language or "en"
        text = WELCOME_TEXT.get(lang, WELCOME_TEXT["en"])

        dlg = WelcomeDialog(self, text, lang)
        dlg.ShowModal()
        if dlg.dont_show_chk.IsChecked():
            self.cfg.show_welcome = False
            save_config(self.cfg)
        dlg.Destroy()

    def _days_names(self):
        return self.strings.get("days_list",["Senin","Selasa","Rabu","Kamis","Jumat","Sabtu","Minggu"])

    def _header_subtitle_text(self):
        if self.cfg.language == "id":
            return "Software pemutaran bell sekolah otomatis karya Tunanetra anak Bangsa."
        return "Automatic school bell playback software created by an Indonesian blind developer."

    def _apply_school_theme(self):
        # Warna tombol mengikuti nuansa sekolah agar tampilan tidak polos.
        neutral_bg = wx.Colour(245, 228, 176)
        neutral_fg = wx.Colour(49, 55, 35)
        action_font = wx.Font(9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        for b in (self.btn_add, self.btn_set, self.btn_log, self.btn_about, self.btn_exit, self.btn_edit, self.btn_delete):
            try:
                b.SetBackgroundColour(neutral_bg)
                b.SetForegroundColour(neutral_fg)
                b.SetMinSize((142, 34))
                b.SetFont(action_font)
            except Exception:
                pass
        try:
            self.btn_activate.SetBackgroundColour(wx.Colour(34, 139, 34))
            self.btn_activate.SetForegroundColour(wx.Colour(255, 255, 255))
            self.btn_activate.SetMinSize((142, 34))
            self.btn_activate.SetFont(action_font)
        except Exception:
            pass
        try:
            self.btn_deactivate.SetBackgroundColour(wx.Colour(178, 34, 34))
            self.btn_deactivate.SetForegroundColour(wx.Colour(255, 255, 255))
            self.btn_deactivate.SetMinSize((142, 34))
            self.btn_deactivate.SetFont(action_font)
        except Exception:
            pass
        try:
            self.Refresh()
        except Exception:
            pass

    def _days_label(self, days):
        try:
            names=self._days_names()
            return ", ".join(names[d] for d in days)
        except Exception:
            return ""

    def _format_last_played(self, value: str):
        if not value:
            return self.strings.get("last_played_empty", "-")
        try:
            stamp = dt.datetime.fromisoformat(value)
            if self.cfg.language == "id":
                return stamp.strftime("%d-%m-%Y %H:%M")
            return stamp.strftime("%Y-%m-%d %H:%M")
        except Exception:
            return value

    def _mark_last_played(self, entry_id: str, played_at_iso: str):
        for e in (self.cfg.entries or []):
            if e.id == entry_id:
                e.last_played = played_at_iso
                break
        save_config(self.cfg)
        wx.CallAfter(self._refresh_list)

    def _apply_button_icons(self):
        size=(18, 18)
        icon_dir = self._asset_icon_dir()

        def _set_btn_icon(button, svg_name, fallback_art):
            if svg_name:
                svg_path = icon_dir / svg_name
                if svg_path.exists():
                    try:
                        svg_data = svg_path.read_bytes()
                        bundle = wx.BitmapBundle.FromSVG(svg_data, wx.Size(*size))
                        bmp = bundle.GetBitmap(wx.Size(*size))
                        button.SetBitmap(bmp)
                        return
                    except Exception:
                        pass
            try:
                button.SetBitmap(wx.ArtProvider.GetBitmap(fallback_art, wx.ART_BUTTON, size))
            except Exception:
                pass

        _set_btn_icon(self.btn_add, "add_bell_plus.svg", wx.ART_PLUS)
        _set_btn_icon(self.btn_set, "settings_gear.svg", wx.ART_EXECUTABLE_FILE)
        _set_btn_icon(self.btn_log, "log_book.svg", wx.ART_REPORT_VIEW)
        _set_btn_icon(self.btn_about, None, wx.ART_INFORMATION)
        _set_btn_icon(self.btn_exit, "exit_door.svg", wx.ART_QUIT)
        _set_btn_icon(self.btn_activate, "activate_check.svg", wx.ART_TICK_MARK)
        _set_btn_icon(self.btn_deactivate, "deactivate_cross.svg", wx.ART_CROSS_MARK)
        _set_btn_icon(self.btn_edit, "edit_pencil.svg", wx.ART_EDIT)
        _set_btn_icon(self.btn_delete, "delete_bin.svg", wx.ART_DELETE)

    def _build_brand_icon(self, size=(32,32)):
        icon = wx.Icon()
        icon_path = self._asset_icon_dir() / "driverbell_brand.svg"
        if icon_path.exists():
            try:
                svg_data = icon_path.read_bytes()
                bundle = wx.BitmapBundle.FromSVG(svg_data, wx.Size(*size))
                bmp = bundle.GetBitmap(wx.Size(*size))
                if bmp.IsOk():
                    icon.CopyFromBitmap(bmp)
                    return icon
            except Exception:
                pass
        try:
            fallback = wx.ArtProvider.GetBitmap(wx.ART_INFORMATION, wx.ART_OTHER, size)
            if fallback.IsOk():
                icon.CopyFromBitmap(fallback)
        except Exception:
            pass
        return icon

    def _asset_icon_dir(self):
        candidates = []
        if hasattr(sys, "_MEIPASS"):
            candidates.append(Path(sys._MEIPASS) / "assets" / "icons")
        candidates.append(Path(sys.executable).resolve().parent / "assets" / "icons")
        candidates.append(Path(__file__).resolve().parent / "assets" / "icons")
        for c in candidates:
            try:
                if c.exists():
                    return c
            except Exception:
                pass
        return candidates[-1]

    def _refresh_list(self):
        prev_name=None
        sel=self._selected_entry()
        if sel: prev_name=sel[0].name
        self.list.DeleteAllItems()
        widths=[240,90,90,220,90,170]
        for i,w in enumerate(widths):
            try: self.list.SetColumnWidth(i,w)
            except Exception: pass
        for e in (self.cfg.entries or []):
            idx=self.list.InsertItem(self.list.GetItemCount(), e.name)
            self.list.SetItem(idx,1,e.start_time)
            self.list.SetItem(idx,2,e.end_time or "")
            self.list.SetItem(idx,3,self._days_label(e.days))
            self.list.SetItem(idx,4,(self.strings["status_active"] if e.active else self.strings["status_inactive"]))
            self.list.SetItem(idx,5,self._format_last_played(e.last_played))
            self.list.SetItemTextColour(idx, wx.Colour(18, 133, 46) if e.active else wx.Colour(180, 22, 22))
        if self.list.GetItemCount()>0:
            target=0
            if prev_name is not None:
                for i in range(self.list.GetItemCount()):
                    if self.list.GetItemText(i)==prev_name:
                        target=i; break
            self._select_index(target, announce=False)

    def _selected_entry(self):
        i=self.list.GetFirstSelected()
        if i==-1: return None
        name=self.list.GetItemText(i)
        for e in (self.cfg.entries or []):
            if e.name==name: return e,i
        return None

    def _entry_status_text(self, e: BellEntry) -> str:
        days_label = self._days_label(e.days)
        status = self.strings["status_active"] if e.active else self.strings["status_inactive"]
        last_played = self._format_last_played(e.last_played)
        if self.cfg.language == "id":
            return (
                f"Dipilih: {e.name}, Mulai {e.start_time}"
                f"{(', Akhir '+e.end_time) if e.end_time else ''}, "
                f"Hari {days_label}, Status {status}, Terakhir {last_played}"
            )
        return (
            f"Selected: {e.name}, Start {e.start_time}"
            f"{(', End '+e.end_time) if e.end_time else ''}, "
            f"Days {days_label}, Status {status}, Last {last_played}"
        )

    def _notify(self,ev,name,ok,details):
        msg=f"{name} {'OK' if ok else 'FAIL'}"
        wx.CallAfter(self.SetStatusText,msg)
        try:
            wx.adv.NotificationMessage(APP_NAME,msg,parent=self).Show(
                timeout=wx.adv.NotificationMessage.Timeout_Auto
            )
        except Exception:
            pass

    def _add(self):
        dlg=EntryDialog(self)
        if dlg.ShowModal()==wx.ID_OK:
            d=dlg.data
            new_e=BellEntry(
                id=str(uuid.uuid4()),
                name=d['name'],
                start_time=d['start_time'],
                end_time=d['end_time'],
                audio_path=d['audio_path'],
                days=d['days'],
                active=True
            )
            self.cfg.entries=(self.cfg.entries or [])+[new_e]
            self.scheduler.clear_jobs_for(new_e.id)
            self.scheduler.schedule_entry(new_e)
            save_config(self.cfg)
            self._refresh_list()
            self.SetStatusText(f"{self.strings['added']} {new_e.name}")
        dlg.Destroy()

    def _edit(self):
        sel=self._selected_entry()
        if not sel: return
        e,idx=sel
        dlg=EntryDialog(self,e)
        if dlg.ShowModal()==wx.ID_OK:
            d=dlg.data
            e.name=d['name']; e.start_time=d['start_time']; e.end_time=d['end_time']; e.audio_path=d['audio_path']; e.days=d['days']
            save_config(self.cfg)
            if e.active:
                self.scheduler.clear_jobs_for(e.id)
                self.scheduler.schedule_entry(e)
            self._refresh_list()
            self.SetStatusText(self.strings['updated'])
        dlg.Destroy()

    def _delete(self):
        sel=self._selected_entry()
        if not sel: return
        e,idx=sel
        if wx.MessageBox(self.strings['confirm_delete'].format(name=e.name),APP_NAME,wx.YES_NO|wx.ICON_WARNING)==wx.YES:
            self.scheduler.clear_jobs_for(e.id)
            self.cfg.entries=[x for x in (self.cfg.entries or []) if x.id!=e.id]
            save_config(self.cfg)
            self._refresh_list()
            self.SetStatusText(self.strings['deleted'])

    def _activate(self):
        sel=self._selected_entry()
        if not sel: return
        e,idx=sel
        if not (e.audio_path or "").strip():
            wx.MessageBox(self.strings['no_audio'],APP_NAME)
            return
        resolved = _resolve_existing_audio_path(e.audio_path)
        if not resolved:
            wx.MessageBox(self.strings['no_audio'],APP_NAME)
            return
        e.audio_path = resolved
        e.active=True
        save_config(self.cfg)
        self.scheduler.clear_jobs_for(e.id)
        self.scheduler.schedule_entry(e)
        self._refresh_list()
        self.SetStatusText(self.strings['activated'])

    def _deactivate(self):
        sel=self._selected_entry()
        if not sel: return
        e,idx=sel
        e.active=False
        save_config(self.cfg)
        self.scheduler.clear_jobs_for(e.id)
        self._refresh_list()
        self.SetStatusText(self.strings['deactivated'])

    def _toggle_active(self):
        sel=self._selected_entry()
        if not sel: return
        e,idx=sel
        if e.active: self._deactivate()
        else: self._activate()

    def _on_show_popup(self,event):
        sel=self._selected_entry()
        if not sel:
            if self.list.GetItemCount()>0: self._select_index(0)
            else: return
        e,idx=self._selected_entry()
        act_label=self.strings['deactivate'] if e.active else self.strings['activate']
        try: self.popup.SetLabel(self.popup_act.GetId(), act_label)
        except Exception: pass
        self.popup.SetLabel(self.popup_edit.GetId(), self.strings['edit'])
        self.popup.SetLabel(self.popup_delete.GetId(), self.strings['delete'])
        self.popup.SetLabel(self.popup_log.GetId(), self.strings['log'])
        self.PopupMenu(self.popup)

    def _on_list_key(self,event):
        code=event.GetKeyCode(); count=self.list.GetItemCount()
        if count==0: event.Skip(); return
        cur=self.list.GetFirstSelected()
        if cur==-1: cur=0
        if self._is_shift_delete(event):
            self._delete()
            return
        if code in (wx.WXK_UP,wx.WXK_DOWN,wx.WXK_PAGEUP,wx.WXK_PAGEDOWN,wx.WXK_HOME,wx.WXK_END):
            if code==wx.WXK_UP: new=max(0,cur-1)
            elif code==wx.WXK_DOWN: new=min(count-1,cur+1)
            elif code==wx.WXK_PAGEUP: new=max(0,cur-10)
            elif code==wx.WXK_PAGEDOWN: new=min(count-1,cur+10)
            elif code==wx.WXK_HOME: new=0
            elif code==wx.WXK_END: new=count-1
            else: new=cur
            self._select_index(new); return
        if code==wx.WXK_APPS or (event.ShiftDown() and code==wx.WXK_F10):
            self._on_show_popup(event); return
        event.Skip()

    def _on_list_char_hook(self,event):
        code=event.GetKeyCode()
        if self._is_shift_delete(event):
            self._delete()
            return
        if code==wx.WXK_TAB:
            self._move_focus_from_list(event.ShiftDown())
            return
        event.Skip()

    def _on_frame_char_hook(self,event):
        # Fallback global: tangani Tab hanya saat fokus memang berada di list.
        try:
            focused = self.FindFocus()
        except Exception:
            focused = None
        if focused == self.list and self._is_shift_delete(event):
            self._delete()
            return
        if focused == self.list and event.GetKeyCode() == wx.WXK_TAB:
            self._move_focus_from_list(event.ShiftDown())
            return
        event.Skip()

    def _is_shift_delete(self, event):
        code = event.GetKeyCode()
        delete_keys = [wx.WXK_DELETE]
        numpad_delete = getattr(wx, "WXK_NUMPAD_DELETE", None)
        if numpad_delete is not None:
            delete_keys.append(numpad_delete)
        return event.ShiftDown() and code in delete_keys

    def _move_focus_from_list(self, backward: bool):
        # Arah fokus eksplisit: cegah singgah ke container "Panel".
        if backward:
            self.btn_delete.SetFocus()
        else:
            self.btn_add.SetFocus()

    def _apply_tab_order(self):
        # Tetapkan urutan tab yang stabil agar list selalu terjangkau pembaca layar.
        try:
            self.btn_set.MoveAfterInTabOrder(self.btn_add)
            self.btn_log.MoveAfterInTabOrder(self.btn_set)
            self.btn_about.MoveAfterInTabOrder(self.btn_log)
            self.btn_exit.MoveAfterInTabOrder(self.btn_about)
            self.btn_activate.MoveAfterInTabOrder(self.btn_exit)
            self.btn_deactivate.MoveAfterInTabOrder(self.btn_activate)
            self.btn_edit.MoveAfterInTabOrder(self.btn_deactivate)
            self.btn_delete.MoveAfterInTabOrder(self.btn_edit)
            self.list.MoveAfterInTabOrder(self.btn_delete)
        except Exception:
            pass

    def _on_list_focus(self,event):
        if self.list.GetItemCount()>0 and self.list.GetFirstSelected()==-1:
            self._select_index(0)
        else:
            sel = self._selected_entry()
            if sel:
                self.SetStatusText(self._entry_status_text(sel[0]))
        event.Skip()

    def _on_list_selected(self,event):
        sel=self._selected_entry()
        if sel:
            self.SetStatusText(self._entry_status_text(sel[0]))
        event.Skip()

    def _on_list_focused(self,event):
        i=event.GetIndex()
        if i!=-1: self._select_index(i)
        event.Skip()

    def _select_index(self,idx,announce=True):
        for i in range(self.list.GetItemCount()):
            self.list.Select(i,on=False)
        self.list.Select(idx,on=True); self.list.Focus(idx); self.list.EnsureVisible(idx)
        if announce and 0<=idx<len(self.cfg.entries or []):
            self.SetStatusText(self._entry_status_text(self.cfg.entries[idx]))

    def _open_settings(self):
        dlg=SettingsDialog(self,self.cfg,self.strings)
        if dlg.ShowModal()==wx.ID_OK:
            save_config(self.cfg)
            # update strings dan UI segera
            self.strings=LANG_STRINGS.get(self.cfg.language,LANG_STRINGS["id"])
            self._refresh_ui_labels()
            self._refresh_list()
            try: self.scheduler.shutdown()
            except Exception: pass
            self.scheduler=BellScheduler(
                self.cfg.timezone,
                self._notify,
                self._mark_last_played,
                lambda: getattr(self.cfg, "output_device", "")
            )
            for e in (self.cfg.entries or []):
                if e.active: self.scheduler.schedule_entry(e)
            try: set_autostart(self.cfg.autostart)
            except Exception: pass
            self.SetStatusText(self.strings["settings_saved"])
        dlg.Destroy()

    def _open_log(self):
        dlg=LogViewer(self); dlg.ShowModal(); dlg.Destroy()

    def _show_about(self):
        dlg=AboutDialog(self, self.cfg.language, self.strings)
        dlg.ShowModal()
        dlg.Destroy()

    def _refresh_ui_labels(self):
        s=self.strings
        self.btn_add.SetLabel(s["add"]); self.btn_set.SetLabel(s["settings"]); self.btn_log.SetLabel(s["log"]); self.btn_about.SetLabel(s.get("about","About")); self.btn_exit.SetLabel(s["exit"])
        self.btn_activate.SetLabel(s["activate"]); self.btn_deactivate.SetLabel(s["deactivate"]); self.btn_edit.SetLabel(s["edit"]); self.btn_delete.SetLabel(s["delete"])
        # rebuild columns
        try:
            while self.list.GetColumnCount()>0: self.list.DeleteColumn(0)
        except Exception:
            pass
        cols=[s["col_name"],s["col_start"],s["col_end"],s["col_days"],s["col_active"],s["col_last_played"]]
        for i,c in enumerate(cols):
            self.list.InsertColumn(i,c)
        # popup labels
        try:
            self.popup.SetLabel(self.popup_act.GetId(), s["activate"])
            self.popup.SetLabel(self.popup_edit.GetId(), s["edit"])
            self.popup.SetLabel(self.popup_delete.GetId(), s["delete"])
            self.popup.SetLabel(self.popup_log.GetId(), s["log"])
        except Exception:
            pass
        self.SetStatusText(s["ready"])
        self.header_subtitle.SetLabel(self._header_subtitle_text())
        self._apply_school_theme()
        self._apply_button_icons()
        self._apply_tab_order()
        if getattr(self,"tray_icon",None):
            try:
                self.tray_icon.SetIcon(self.tray_icon.GetIcon(), s.get("tray_title",APP_NAME))
            except Exception:
                pass

    # --- PATCH: handle EVT_CLOSE untuk run_in_background ---
    def _on_close(self, event):
        if self.cfg.run_in_background and not self._really_quit:
            # sembunyikan ke tray, Veto agar frame tidak Destroy
            self._show_in_tray()
            self.Hide()
            self.SetStatusText(self.strings.get("tray_title", APP_NAME))
            if event: event.Veto()
            return
        # benar-benar keluar
        try:
            if getattr(self,"tray_icon",None):
                self.tray_icon.RemoveIcon(); self.tray_icon.Destroy()
        except Exception:
            pass
        try:
            self.scheduler.shutdown()
        except Exception:
            pass
        self.Destroy()

    # Catatan: kita biarkan override Close() lama tidak dipakai lagi. Semua exit diarahkan via _on_close.

    def _show_in_tray(self):
        if getattr(self,"tray_icon",None):
            return
        icon=self._build_brand_icon((16,16))
        self.tray_icon=TaskBarIcon(self,icon)
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DCLICK,lambda e: self._restore_from_tray())

    def _restore_from_tray(self):
        if getattr(self,"tray_icon",None):
            try:
                self.tray_icon.RemoveIcon(); self.tray_icon.Destroy()
            except Exception:
                pass
            self.tray_icon=None
        self.Show(); self.Raise(); self.SetStatusText(self.strings["ready"])

class TaskBarIcon(wx.adv.TaskBarIcon):
    ID_SHOW=wx.NewIdRef(); ID_QUIT=wx.NewIdRef()
    def __init__(self,frame,icon):
        super().__init__()
        self.frame=frame
        self.SetIcon(icon, frame.strings.get("tray_title",APP_NAME))
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DCLICK,self.on_restore)
    def CreatePopupMenu(self):
        menu=wx.Menu()
        menu.Append(self.ID_SHOW,self.frame.strings.get("show","Show"))
        menu.AppendSeparator()
        menu.Append(self.ID_QUIT,self.frame.strings.get("quit","Quit"))
        self.Bind(wx.EVT_MENU, lambda e: self.on_restore(e) if e.GetId()==self.ID_SHOW.GetId() else self.on_quit(e), id=self.ID_SHOW.GetId())
        self.Bind(wx.EVT_MENU, self.on_quit, id=self.ID_QUIT.GetId())
        return menu
    def on_restore(self,event): self.frame._restore_from_tray()
    def on_quit(self,event):
        # PATCH: set flag agar _on_close keluar total
        self.frame._really_quit = True
        self.frame.Close()

class EntryDialog(wx.Dialog):
    def __init__(self,parent,entry:BellEntry=None):
        self.parent=parent; self.strings=parent.strings
        title=self.strings["edit"] if entry else self.strings["add"]
        super().__init__(parent,title=title,size=(560,320))
        pnl=wx.Panel(self); v=wx.BoxSizer(wx.VERTICAL)
        grid=wx.FlexGridSizer(0,4,8,8); grid.AddGrowableCol(1,1)

        grid.Add(wx.StaticText(pnl,label=self.strings["entry_name"]))
        self.t_name=wx.TextCtrl(pnl); grid.Add(self.t_name,1,wx.EXPAND); grid.Add((1,1)); grid.Add((1,1))

        grid.Add(wx.StaticText(pnl,label=self.strings["start_time"]))
        self.t_start_h=wx.SpinCtrl(pnl,min=0,max=23,initial=7)
        self.t_start_m=wx.SpinCtrl(pnl,min=0,max=59,initial=0)
        h1=wx.BoxSizer(wx.HORIZONTAL); h1.Add(self.t_start_h,0); h1.Add(wx.StaticText(pnl,label=":")); h1.Add(self.t_start_m,0)
        grid.Add(h1,0,wx.EXPAND); grid.Add((1,1)); grid.Add((1,1))

        grid.Add(wx.StaticText(pnl,label=self.strings.get("end_time",self.strings.get("end_time","Jam Akhir (opsional):"))))
        self.t_end_h=wx.SpinCtrl(pnl,min=0,max=23)
        self.t_end_m=wx.SpinCtrl(pnl,min=0,max=59)
        h2=wx.BoxSizer(wx.HORIZONTAL); h2.Add(self.t_end_h,0); h2.Add(wx.StaticText(pnl,label=":")); h2.Add(self.t_end_m,0)
        grid.Add(h2,0,wx.EXPAND); grid.Add((1,1)); grid.Add((1,1))

        grid.Add(wx.StaticText(pnl,label=self.strings["audio_file"]))
        h_audio=wx.BoxSizer(wx.HORIZONTAL); self.t_audio=wx.TextCtrl(pnl); btn_browse=wx.Button(pnl,label=self.strings["browse"]); self.btn_preview=wx.Button(pnl,label=self.strings.get("preview_audio","Preview Audio"))
        h_audio.add = h_audio.Add  # alias helper
        h_audio.add(self.t_audio,1,wx.RIGHT,6); h_audio.add(btn_browse,0,wx.RIGHT,6); h_audio.add(self.btn_preview,0)
        grid.Add(h_audio,1,wx.EXPAND); grid.Add((1,1)); grid.Add((1,1))

        grid.Add(wx.StaticText(pnl,label=self.strings["days_label"]))
        h_days=wx.BoxSizer(wx.HORIZONTAL); self.cb_days=[]
        for i,day in enumerate(self.strings.get("days_list",["Senin","Selasa","Rabu","Kamis","Jumat","Sabtu","Minggu"])):
            cb=wx.CheckBox(pnl,label=day); h_days.Add(cb,0,wx.RIGHT,6); self.cb_days.append(cb)
        grid.Add(h_days,1,wx.EXPAND); grid.Add((1,1)); grid.Add((1,1))

        v.Add(grid,1,wx.ALL|wx.EXPAND,12)
        v.Add(wx.StaticLine(pnl),0,wx.EXPAND|wx.LEFT|wx.RIGHT,12)
        h=wx.BoxSizer(wx.HORIZONTAL)
        btn_ok=wx.Button(pnl,wx.ID_OK,label=self.strings.get("save","Save"))
        btn_cancel=wx.Button(pnl,wx.ID_CANCEL,label=self.strings.get("cancel","Cancel"))
        h.Add(btn_cancel,0,wx.ALL,8); h.Add(btn_ok,0,wx.ALL,8)
        v.Add(h,0,wx.ALIGN_RIGHT)
        pnl.SetSizer(v)

        btn_browse.Bind(wx.EVT_BUTTON,self._browse)
        self.btn_preview.Bind(wx.EVT_BUTTON,self._toggle_preview_audio)
        self.Bind(wx.EVT_BUTTON,self._on_ok,btn_ok)
        self.Bind(wx.EVT_BUTTON,self._on_cancel,btn_cancel)
        self.Bind(wx.EVT_CHAR_HOOK,self._on_char_hook)
        self.Bind(wx.EVT_CLOSE,self._on_close)
        self.preview_player = None
        self.preview_is_playing = False
        self.preview_timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER,self._on_preview_timer,self.preview_timer)
        self.data=None

        if entry:
            self.t_name.SetValue(entry.name)
            try:
                sh,sm=map(int,entry.start_time.split(":"))
                self.t_start_h.SetValue(sh); self.t_start_m.SetValue(sm)
            except Exception:
                pass
            if entry.end_time:
                try:
                    eh,em=map(int,entry.end_time.split(":"))
                    self.t_end_h.SetValue(eh); self.t_end_m.SetValue(em)
                except Exception:
                    pass
            self.t_audio.SetValue(entry.audio_path)
            for i in entry.days:
                if 0<=i<len(self.cb_days): self.cb_days[i].SetValue(True)

    def _browse(self,event):
        with wx.FileDialog(self,self.strings["browse"],
                           wildcard="Supported audio (*.wav;*.mp3;*.aac;*.ogg;*.flac;*.m4a)|*.wav;*.mp3;*.aac;*.ogg;*.flac;*.m4a|All files (*.*)|*.*",
                           style=wx.FD_OPEN|wx.FD_FILE_MUST_EXIST) as fd:
            if fd.ShowModal()==wx.ID_OK:
                chosen = _normalize_audio_path(fd.GetPath())
                self.t_audio.SetValue(chosen)

    def _set_preview_state(self, is_playing: bool):
        self.preview_is_playing = bool(is_playing)
        self.btn_preview.SetLabel(
            self.strings.get("stop_preview","Stop Preview")
            if self.preview_is_playing
            else self.strings.get("preview_audio","Preview Audio")
        )
        if self.preview_is_playing:
            if not self.preview_timer.IsRunning():
                self.preview_timer.Start(500)
        else:
            if self.preview_timer.IsRunning():
                self.preview_timer.Stop()

    def _toggle_preview_audio(self, event):
        if self.preview_is_playing:
            self._stop_preview_now()
            return
        path_input = self.t_audio.GetValue().strip()
        if not path_input:
            wx.MessageBox(self.strings["no_audio"],APP_NAME); return
        try:
            self.preview_player = AudioPlayer(getattr(self.parent.cfg, "output_device", ""))
            ok = self.preview_player.play_loop(path_input)
            if not ok:
                wx.MessageBox(self.strings.get("audio_pick","Please choose a valid audio file."),APP_NAME); return
            self._set_preview_state(True)
        except Exception:
            wx.MessageBox(self.strings.get("audio_pick","Please choose a valid audio file."),APP_NAME); return

    def _stop_preview_now(self):
        try:
            if self.preview_player:
                self.preview_player.stop()
        except Exception:
            pass
        self._set_preview_state(False)

    def _on_preview_timer(self, event):
        if self.preview_player and self.preview_player.is_playing():
            return
        self._set_preview_state(False)

    def _on_cancel(self, event):
        self._stop_preview_now()
        self.EndModal(wx.ID_CANCEL)

    def _on_char_hook(self, event):
        if event.GetKeyCode() == wx.WXK_ESCAPE:
            self._stop_preview_now()
        event.Skip()

    def _on_close(self, event):
        self._stop_preview_now()
        event.Skip()

    def _on_ok(self,event):
        self._stop_preview_now()
        name=self.t_name.GetValue().strip()
        if not name:
            wx.MessageBox(self.strings["name_required"],APP_NAME); return
        sh,sm=self.t_start_h.GetValue(),self.t_start_m.GetValue()
        eh,em=self.t_end_h.GetValue(),self.t_end_m.GetValue()
        start_time=f"{sh:02d}:{sm:02d}"
        end_time=f"{eh:02d}:{em:02d}" if (eh or em) else ""
        days=[i for i,cb in enumerate(self.cb_days) if cb.GetValue()]
        if not days:
            wx.MessageBox(self.strings["select_day"],APP_NAME); return
        audio_path=_normalize_audio_path(self.t_audio.GetValue().strip())
        if not audio_path:
            wx.MessageBox(self.strings["no_audio"],APP_NAME); return
        resolved = _resolve_existing_audio_path(audio_path)
        if not resolved:
            wx.MessageBox(self.strings["no_audio"],APP_NAME); return
        self.data={"name":name,"start_time":start_time,"end_time":end_time,"audio_path":resolved,"days":days}
        self.EndModal(wx.ID_OK)

class SettingsDialog(wx.Dialog):
    # display label, saved value (label, iana)
    TIMEZONE_OPTIONS = [
        ("WIB (Asia/Jakarta)", "Asia/Jakarta"),
        ("WITA (Asia/Makassar)", "Asia/Makassar"),
        ("WIT (Asia/Jayapura)", "Asia/Jayapura"),
    ]
    def __init__(self,parent,cfg:AppConfig,strings):
        title=strings.get("settings","Settings")
        super().__init__(parent,title=title,size=(560,430))
        self.cfg=cfg; self.strings=strings
        pnl=wx.Panel(self); v=wx.BoxSizer(wx.VERTICAL)

        v.Add(wx.StaticText(pnl,label=self.strings.get("timezone","Timezone")+":"),0,wx.ALL,8)
        tz_labels=[lbl for lbl,iana in self.TIMEZONE_OPTIONS]
        self.combo_tz=wx.ComboBox(pnl,choices=tz_labels,style=wx.CB_READONLY)
        cur=cfg.timezone or DEFAULT_TZ; sel_idx=0
        for i,(lbl,iana) in enumerate(self.TIMEZONE_OPTIONS):
            if iana==cur: sel_idx=i; break
        self.combo_tz.SetSelection(sel_idx)
        v.Add(self.combo_tz,0,wx.EXPAND|wx.LEFT|wx.RIGHT,8)

        v.Add(wx.StaticText(pnl,label=self.strings.get("output_device","Output Device")+":"),0,wx.ALL,8)
        self.output_label_default = self.strings.get("output_device_default","System Default")
        detected_devices = list_audio_output_devices()
        self.output_choices = [self.output_label_default] + detected_devices
        self.combo_output=wx.ComboBox(pnl,choices=self.output_choices,style=wx.CB_READONLY)
        saved_device = (getattr(cfg, "output_device", "") or "").strip()
        if saved_device and saved_device in self.output_choices:
            self.combo_output.SetSelection(self.output_choices.index(saved_device))
        else:
            self.combo_output.SetSelection(0)
        v.Add(self.combo_output,0,wx.EXPAND|wx.LEFT|wx.RIGHT,8)

        v.Add(wx.StaticText(pnl,label=self.strings.get("language","Language")+":"),0,wx.ALL,8)
        self.rb_lang=wx.RadioBox(pnl,choices=["Indonesia","English"],majorDimension=2,style=wx.RA_SPECIFY_COLS)
        self.rb_lang.SetSelection(0 if cfg.language=="id" else 1)
        v.Add(self.rb_lang,0,wx.EXPAND|wx.LEFT|wx.RIGHT,8)

        self.cb_background=wx.CheckBox(pnl,label=self.strings.get("run_in_background"))
        self.cb_background.SetValue(bool(cfg.run_in_background))
        v.Add(self.cb_background,0,wx.ALL|wx.EXPAND,8)

        self.cb_autostart=wx.CheckBox(pnl,label=self.strings.get("autostart"))
        self.cb_autostart.SetValue(bool(cfg.autostart))
        v.Add(self.cb_autostart,0,wx.ALL|wx.EXPAND,8)

        self.cb_show_welcome=wx.CheckBox(pnl,label=self.strings.get("show_welcome_on_startup","Show welcome dialog on startup"))
        self.cb_show_welcome.SetValue(bool(getattr(cfg,"show_welcome",True)))
        v.Add(self.cb_show_welcome,0,wx.ALL|wx.EXPAND,8)

        v.Add(wx.StaticLine(pnl),0,wx.EXPAND|wx.LEFT|wx.RIGHT,8)
        h=wx.BoxSizer(wx.HORIZONTAL)
        btn_ok=wx.Button(pnl,wx.ID_OK,label=self.strings.get("save","Save"))
        btn_cancel=wx.Button(pnl,wx.ID_CANCEL,label=self.strings.get("cancel","Cancel"))
        h.Add(btn_cancel,0,wx.ALL,8); h.Add(btn_ok,0,wx.ALL,8)
        v.Add(h,0,wx.ALIGN_RIGHT)
        pnl.SetSizer(v)

        self.Bind(wx.EVT_BUTTON,self._on_ok,btn_ok)

    def _on_ok(self,event):
        sel=self.combo_tz.GetSelection()
        if sel<0: sel=0
        self.cfg.timezone=self.TIMEZONE_OPTIONS[sel][1]
        self.cfg.language="id" if self.rb_lang.GetSelection()==0 else "en"
        self.cfg.run_in_background=bool(self.cb_background.GetValue())
        self.cfg.autostart=bool(self.cb_autostart.GetValue())
        self.cfg.show_welcome=bool(self.cb_show_welcome.GetValue())
        out_sel=self.combo_output.GetSelection()
        self.cfg.output_device="" if out_sel<=0 else self.output_choices[out_sel]
        self.EndModal(wx.ID_OK)

class LogViewer(wx.Dialog):
    def __init__(self,parent):
        super().__init__(parent,title=parent.strings.get("log","Log"),size=(780,460))
        pnl=wx.Panel(self); v=wx.BoxSizer(wx.VERTICAL)
        btn_refresh=wx.Button(pnl,label=parent.strings.get("updated_label","Refresh"))
        v.Add(btn_refresh,0,wx.ALL,8)
        self.txt=wx.TextCtrl(pnl,style=wx.TE_MULTILINE|wx.TE_READONLY)
        v.Add(self.txt,1,wx.EXPAND|wx.ALL,8)
        pnl.SetSizer(v)
        btn_refresh.Bind(wx.EVT_BUTTON,lambda e: self._load(parent))
        self._load(parent)
    def _load(self, parent):
        text = LOG_PATH.read_text(encoding="utf-8", errors="ignore") if LOG_PATH.exists() else (
            "No logs." if parent.cfg.language == "en" else "Belum ada log."
        )
        self.txt.SetValue(text)
class AboutDialog(wx.Dialog):
    def __init__(self, parent, lang="id", strings=None):
        if lang not in ("id", "en"):
            lang = "en"
        strings = strings or LANG_STRINGS.get(lang, LANG_STRINGS["en"])

        super().__init__(parent, title=strings.get("about", "About"), size=(700, 470))

        pnl=wx.Panel(self)
        v=wx.BoxSizer(wx.VERTICAL)

        txt=wx.StaticText(pnl,label=strings.get("about_text",""))
        txt.Wrap(650)
        v.Add(txt,1,wx.EXPAND|wx.ALL,10)

        h=wx.BoxSizer(wx.HORIZONTAL)
        donate_label = "Donasi Sekarang" if lang=="id" else "Donate now"
        close_label = "Tutup" if lang=="id" else "Close"
        btn_donate=wx.Button(pnl,wx.ID_ANY,label=donate_label)
        btn_close=wx.Button(pnl,wx.ID_CANCEL,label=close_label)
        h.Add(btn_donate,0,wx.ALL,6)
        h.Add(btn_close,0,wx.ALL,6)
        v.Add(h,0,wx.ALIGN_RIGHT|wx.RIGHT|wx.BOTTOM,8)
        pnl.SetSizer(v)

        btn_donate.Bind(wx.EVT_BUTTON,lambda e: self._open_donate(lang))

    def _open_donate(self, lang):
        dlg=DonateDialog(self, lang=lang)
        dlg.ShowModal()
        dlg.Destroy()


class DonateDialog(wx.Dialog):
    def __init__(self, parent, lang="id"):
        if lang not in ("id", "en"):
            lang = "en"

        if lang == "id":
            title = "Daftar Donasi DriverBell"
            intro = "Silakan pilih metode donasi berikut:"
            copy_jago = "Salin rekening Bank Jago"
            copy_gopay = "Salin akun GoPay"
            copy_paypal = "Salin rekening PayPal"
            close_label = "Tutup"
            jago_block = "Bank Jago:\nNomor rekening: 506727819682\nAtas nama: Dwi Citolaksono"
            gopay_block = "GoPay\nNomor: 085798020576"
            paypal_block = "PayPal\nEmail: dwicitolaksonoo@gmail.com"
        else:
            title = "DriverBell Donation Accounts"
            intro = "Please choose one of the following donation methods:"
            copy_jago = "Copy the Jago Bank account"
            copy_gopay = "Copy the GoPay account"
            copy_paypal = "Copy the PayPal account"
            close_label = "Close"
            jago_block = "Bank Jago:\nAccount number: 506727819682\nAccount name: Dwi Citolaksono"
            gopay_block = "GoPay\nNumber: 085798020576"
            paypal_block = "PayPal\nEmail: dwicitolaksonoo@gmail.com"

        super().__init__(parent, title=title, size=(560, 560), style=wx.DEFAULT_DIALOG_STYLE | wx.STAY_ON_TOP)

        pnl=wx.Panel(self)
        v=wx.BoxSizer(wx.VERTICAL)
        intro_txt=wx.StaticText(pnl,label=intro)
        intro_txt.Wrap(500)
        v.Add(intro_txt,0,wx.ALL|wx.EXPAND,10)

        self.t_jago=wx.TextCtrl(pnl,value=jago_block,style=wx.TE_MULTILINE|wx.TE_READONLY)
        self.t_jago.SetMinSize((-1,80))
        v.Add(self.t_jago,0,wx.LEFT|wx.RIGHT|wx.TOP|wx.EXPAND,10)
        self.btn_copy_jago=wx.Button(pnl,wx.ID_ANY,label=copy_jago)
        v.Add(self.btn_copy_jago,0,wx.LEFT|wx.RIGHT|wx.TOP,10)

        self.t_gopay=wx.TextCtrl(pnl,value=gopay_block,style=wx.TE_MULTILINE|wx.TE_READONLY)
        self.t_gopay.SetMinSize((-1,62))
        v.Add(self.t_gopay,0,wx.LEFT|wx.RIGHT|wx.TOP|wx.EXPAND,10)
        self.btn_copy_gopay=wx.Button(pnl,wx.ID_ANY,label=copy_gopay)
        v.Add(self.btn_copy_gopay,0,wx.LEFT|wx.RIGHT|wx.TOP,10)

        self.t_paypal=wx.TextCtrl(pnl,value=paypal_block,style=wx.TE_MULTILINE|wx.TE_READONLY)
        self.t_paypal.SetMinSize((-1,62))
        v.Add(self.t_paypal,0,wx.LEFT|wx.RIGHT|wx.TOP|wx.EXPAND,10)
        self.btn_copy_paypal=wx.Button(pnl,wx.ID_ANY,label=copy_paypal)
        v.Add(self.btn_copy_paypal,0,wx.LEFT|wx.RIGHT|wx.TOP,10)

        self.btn_close=wx.Button(pnl,wx.ID_CANCEL,label=close_label)
        v.Add(self.btn_close,0,wx.ALIGN_RIGHT|wx.ALL,10)

        pnl.SetSizer(v)

        self.btn_copy_jago.Bind(wx.EVT_BUTTON, lambda e: self._copy_to_clipboard("506727819682"))
        self.btn_copy_gopay.Bind(wx.EVT_BUTTON, lambda e: self._copy_to_clipboard("085798020576"))
        self.btn_copy_paypal.Bind(wx.EVT_BUTTON, lambda e: self._copy_to_clipboard("dwicitolaksonoo@gmail.com"))

    def _copy_to_clipboard(self, value):
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(wx.TextDataObject(value))
            wx.TheClipboard.Close()


class WelcomeDialog(wx.Dialog):
    def __init__(self, parent, text: dict, lang: str):
        super().__init__(
            parent,
            title=text.get("title", "Welcome"),
            style=wx.DEFAULT_DIALOG_STYLE | wx.STAY_ON_TOP,
        )
        self._lang = lang

        message = wx.StaticText(self, label=text.get("message", ""))
        message.Wrap(420)

        self.dont_show_chk = wx.CheckBox(self, label=text.get("dont_show", ""))

        ok_label = text.get("ok", "OK")
        donate_label = text.get("donate", "Donate")

        ok_btn = wx.Button(self, wx.ID_OK, ok_label)
        donate_btn = wx.Button(self, wx.ID_ANY, donate_label)
        donate_btn.SetDefault()

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.Add(ok_btn, 0, wx.ALL, 5)
        btn_sizer.Add(donate_btn, 0, wx.ALL, 5)

        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(message, 0, wx.ALL | wx.EXPAND, 10)
        main_sizer.Add(self.dont_show_chk, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        main_sizer.Add(btn_sizer, 0, wx.ALIGN_RIGHT | wx.ALL, 5)

        self.SetSizerAndFit(main_sizer)

        donate_btn.Bind(wx.EVT_BUTTON, self.on_donate)

    def on_donate(self, event):
        dlg = DonateDialog(self, lang=self._lang)
        dlg.ShowModal()
        dlg.Destroy()

def get_startup_folder():
    return str(os.path.join(os.getenv('APPDATA'), r'Microsoft\\Windows\\Start Menu\\Programs\\Startup'))

def _startup_paths():
    startup = get_startup_folder()
    return {
        "lnk": os.path.join(startup, "DriverBell.lnk"),
        "bat": os.path.join(startup, "DriverBell.bat"),
        "vbs": os.path.join(startup, "DriverBell.vbs"),
    }

def _remove_startup_file(path: str):
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass

def _startup_command(target: str, script: str) -> str:
    if getattr(sys, "frozen", False):
        return f'"{target}"'
    return f'"{target}" "{script}"'

def _write_hidden_vbs_startup(path: str, command: str):
    escaped = command.replace('"', '""')
    with open(path, "w", encoding="utf-8") as f:
        f.write('Set shell = CreateObject("WScript.Shell")\n')
        f.write(f'shell.Run "{escaped}", 0, False\n')

# --- PATCH BESAR: Autostart yang benar (COM init + bedakan frozen/script) ---
def create_shortcut_in_startup(enabled=True):
    startup=get_startup_folder()
    target=sys.executable           # jika frozen -> exe aplikasi; jika script -> python.exe
    script=os.path.abspath(sys.argv[0])
    paths=_startup_paths()
    path=paths["lnk"]
    command=_startup_command(target, script)

    if enabled:
        try:
            _remove_startup_file(paths["bat"])
            _remove_startup_file(paths["vbs"])
            if has_pywin:
                # Inisialisasi COM sebelum pakai IShellLink
                pythoncom.CoInitialize()
                try:
                    shell_link=pythoncom.CoCreateInstance(
                        shellcon.CLSID_ShellLink, None,
                        pythoncom.CLSCTX_INPROC_SERVER, shellcon.IID_IShellLink
                    )
                    if getattr(sys, "frozen", False):
                        # EXE hasil PyInstaller: target = exe; tanpa argumen
                        shell_link.SetPath(target)
                        shell_link.SetArguments("")
                        shell_link.SetWorkingDirectory(os.path.dirname(target))
                    else:
                        # Mode script: target = python.exe; argumen = path script
                        shell_link.SetPath(target)
                        shell_link.SetArguments(f'"{script}"')
                        shell_link.SetWorkingDirectory(os.path.dirname(script))
                    shell_link.SetDescription("DriverBell - Auto start")
                    persist_file=shell_link.QueryInterface(pythoncom.IID_IPersistFile)
                    persist_file.Save(path,0)
                finally:
                    pythoncom.CoUninitialize()
            else:
                # Fallback tanpa CMD: wscript menjalankan aplikasi tersembunyi.
                _remove_startup_file(path)
                _write_hidden_vbs_startup(paths["vbs"], command)
        except Exception:
            # Jika pembuatan .lnk gagal, tetap hindari .bat/CMD.
            _remove_startup_file(path)
            _write_hidden_vbs_startup(paths["vbs"], command)
    else:
        for old_path in paths.values():
            _remove_startup_file(old_path)

def set_autostart(enabled:bool):
    if os.name!='nt': return
    create_shortcut_in_startup(enabled)

def ensure_files():
    _migrate_legacy_files()
    if not CONFIG_PATH.exists():
        # ambil zona lokal sebagai default agar lebih ramah pengguna
        try:
            tz = str(get_localzone()) or DEFAULT_TZ
        except Exception:
            tz = DEFAULT_TZ
        save_config(AppConfig(timezone=tz, entries=[]))
    if not LOG_PATH.exists():
        append_log({"timestamp": dt.datetime.now().isoformat(),
                    "name": "SYSTEM", "action": "INIT", "result": "OK", "details": "Log dibuat"})

def main():
    ensure_files()
    cfg = load_config()

    # Buat aplikasi wx
    app = wx.App(False)
    initialize_audio_backend()

    # Single instance checker:
    # gunakan nama aplikasi dan folder APPDATA sebagai identitas
    checker = wx.SingleInstanceChecker(APP_NAME, str(APPDATA_DIR))
    if checker.IsAnotherRunning():
        msg = (
            "DriverBell is already running.\nPlease check the DriverBell icon in the System Tray or Taskbar."
            if cfg.language == "en"
            else "DriverBell sudah berjalan.\nSilakan cek ikon DriverBell di System Tray atau Taskbar."
        )
        wx.MessageBox(
            msg,
            APP_NAME,
            wx.OK | wx.ICON_INFORMATION
        )
        return

    frame = MainFrame(cfg)
    try:
        set_autostart(cfg.autostart)
    except Exception:
        # kalau gagal autostart, biarkan saja, tidak mengganggu bel
        pass

    app.MainLoop()

if __name__=="__main__":
    main()


