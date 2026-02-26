import os
import sys

import subprocess
import importlib.util
import threading
import random
import re
import json
import time
import shutil
import zipfile
from pathlib import Path
from tkinter import messagebox, filedialog
from io import BytesIO
from urllib.request import urlopen, Request

# ============================================================================
# PYINSTALLER
# ============================================================================

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def _python_exe():
    """Return the real Python interpreter path, not the frozen exe."""
    if getattr(sys, 'frozen', False):
        # Frozen exe — find system Python
        import shutil as _sh
        p = _sh.which("python") or _sh.which("python3")
        if p:
            return p
        # Fallback: check common install locations
        for candidate in [
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Python", "Python311", "python.exe"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Python", "Python312", "python.exe"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Python", "Python310", "python.exe"),
            r"C:\Python311\python.exe",
            r"C:\Python312\python.exe",
        ]:
            if os.path.isfile(candidate):
                return candidate
        return "python"  # Hope it's on PATH
    return sys.executable

# Ensure CustomTkinter (skip when running as frozen exe — already bundled)
if not getattr(sys, 'frozen', False) and importlib.util.find_spec("customtkinter") is None:
    subprocess.run([_python_exe(), "-m", "pip", "install", "customtkinter", "-q"])

import customtkinter as ctk
from PIL import Image, ImageDraw


# ============================================================================
# SETTINGS PERSISTENCE
# ============================================================================

_SETTINGS_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "RLA_PlaylistCreator"
_SETTINGS_PATH = _SETTINGS_DIR / "settings.json"


def _default_output_folder():
    """Return BeamNG mods folder if it exists, otherwise Desktop"""
    beamng = Path(os.environ.get("LOCALAPPDATA", "")) / "BeamNG" / "BeamNG.drive" / "current" / "mods"
    if beamng.is_dir():
        return str(beamng)
    return str(Path.home() / "Desktop")


def _default_settings():
    return {
        "output_folder": _default_output_folder(),
        "audio_format": "mp3",
        "audio_quality": "192",
        "appearance": "System",
        "last_update_check": 0,
    }


def load_settings():
    defaults = _default_settings()
    try:
        if _SETTINGS_PATH.is_file():
            with open(_SETTINGS_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            # Merge: saved values override defaults, but missing keys get defaults
            for k, v in defaults.items():
                saved.setdefault(k, v)
            return saved
    except Exception:
        pass
    return defaults


def save_settings(settings):
    try:
        _SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        with open(_SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except Exception:
        pass


# ============================================================================
# BACKEND CONSTANTS
# ============================================================================

__version__ = "2.0.0"
_GITHUB_REPO = "RedlineAuto/PlaylistCreator"
_UPDATE_COOLDOWN = 3600  # seconds between update checks

FFMPEG_DIR = _SETTINGS_DIR / "ffmpeg"
FFMPEG_URL = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip"
_ffmpeg_path = None


# ============================================================================
# AUTO-UPDATE
# ============================================================================

def _parse_version(tag):
    """Parse 'v2.1.0' or '2.1.0' into a tuple for comparison."""
    tag = tag.lstrip("vV")
    return tuple(int(p) for p in tag.split(".") if p.isdigit())


def _should_check_update():
    """Return True if enough time has passed since last check."""
    settings = load_settings()
    last_check = settings.get("last_update_check", 0)
    return (time.time() - last_check) >= _UPDATE_COOLDOWN


def _mark_update_checked():
    """Record that we just checked for updates."""
    settings = load_settings()
    settings["last_update_check"] = time.time()
    save_settings(settings)


def check_for_update(log_callback=None):
    """Check GitHub for a newer release.
    Returns dict with version/download_url/asset_name, or None."""
    try:
        url = f"https://api.github.com/repos/{_GITHUB_REPO}/releases/latest"
        req = Request(url, headers={
            'User-Agent': f'BeamNG-PlaylistCreator/{__version__}',
            'Accept': 'application/vnd.github.v3+json',
        })
        response = urlopen(req, timeout=10)
        data = json.loads(response.read().decode('utf-8'))

        latest_tag = data.get("tag_name", "")
        latest_ver = _parse_version(latest_tag)
        current_ver = _parse_version(__version__)

        if latest_ver <= current_ver:
            if log_callback:
                log_callback(f"App is up to date (v{__version__})")
            return None

        # Find installer asset matching our brand
        brand = _SETTINGS_DIR.name
        download_url = None
        asset_name = None
        for asset in data.get("assets", []):
            name = asset.get("name", "")
            if name.lower().endswith(".exe") and brand.lower() in name.lower():
                download_url = asset["browser_download_url"]
                asset_name = name
                break

        if not download_url:
            if log_callback:
                log_callback(f"Update {latest_tag} found but no installer for {brand}")
            return None

        if log_callback:
            log_callback(f"Update available: v{__version__} -> {latest_tag}")

        return {
            "version": latest_tag,
            "download_url": download_url,
            "asset_name": asset_name,
        }

    except Exception as e:
        if log_callback:
            log_callback(f"Update check failed: {str(e)[:120]}")
        return None


def download_update(download_url, asset_name, log_callback=None, stop_check=None):
    """Download installer to _SETTINGS_DIR. Returns path on success, None on failure."""
    installer_path = _SETTINGS_DIR / asset_name

    try:
        _SETTINGS_DIR.mkdir(parents=True, exist_ok=True)

        if log_callback:
            log_callback(f"Downloading update: {asset_name}...")

        req = Request(download_url, headers={
            'User-Agent': f'BeamNG-PlaylistCreator/{__version__}'
        })
        response = urlopen(req, timeout=60)
        total = int(response.headers.get('Content-Length', 0))
        downloaded = 0

        with open(installer_path, 'wb') as f:
            while True:
                if stop_check and stop_check():
                    installer_path.unlink(missing_ok=True)
                    return None
                chunk = response.read(256 * 1024)
                if not chunk:
                    break
                f.write(chunk)
                downloaded += len(chunk)
                if log_callback and total > 0:
                    mb_done = downloaded / (1024 * 1024)
                    mb_total = total / (1024 * 1024)
                    pct = int((downloaded / total) * 100)
                    log_callback(f"Downloading update... {mb_done:.1f}/{mb_total:.1f} MB ({pct}%)")

        if installer_path.exists() and installer_path.stat().st_size > 0:
            if log_callback:
                log_callback("Download complete")
            return installer_path

        installer_path.unlink(missing_ok=True)
        return None

    except Exception as e:
        installer_path.unlink(missing_ok=True)
        if log_callback:
            log_callback(f"Update download failed: {str(e)[:120]}")
        return None


# ============================================================================
# BACKEND UTILITIES
# ============================================================================

def _hidden_startupinfo():
    """Return startupinfo to hide console windows on Windows"""
    if sys.platform == 'win32':
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = subprocess.SW_HIDE
        return si
    return None


def _strip_ansi(text):
    """Remove ANSI escape codes from text"""
    return re.sub(r'\x1b\[[0-9;]*m', '', text)


def clean_filename(filename):
    """Clean filename to be filesystem safe"""
    invalid_chars = '<>:"|?*\\/\'.[]{}()+-!@#$%^&=`~'
    cleaned = ''.join(c for c in filename if c not in invalid_chars)
    cleaned = cleaned.replace(' ', '_')
    return cleaned.strip()


def find_available_name(base_name, parent_dir, suffix="_Radiopack", start_version=1, log_callback=None):
    """Find an available name with auto-versioning."""
    version_pattern = r'^(.+?)_v(\d+)$'
    match = re.match(version_pattern, base_name)

    if match:
        real_base_name = match.group(1)
        input_version = int(match.group(2))
        target = parent_dir / f"{base_name}{suffix}"
        if not target.exists():
            return base_name, target
        version = input_version + 1
    else:
        real_base_name = base_name
        target = parent_dir / f"{base_name}{suffix}"
        if not target.exists():
            return base_name, target
        version = start_version

    while True:
        versioned_name = f"{real_base_name}_v{version}"
        target = parent_dir / f"{versioned_name}{suffix}"
        if not target.exists():
            if log_callback:
                log_callback(f"'{base_name}{suffix}' already exists -> using: {versioned_name}")
            return versioned_name, target
        version += 1
        if version > 999:
            raise Exception(f"Too many versions of {real_base_name}")


# ============================================================================
# FFMPEG MANAGEMENT
# ============================================================================

def get_ffmpeg_path():
    """Find ffmpeg. Checks: bundled with exe -> AppData -> system PATH"""
    global _ffmpeg_path
    if _ffmpeg_path and Path(_ffmpeg_path).exists():
        return _ffmpeg_path

    # Bundled with PyInstaller exe
    if getattr(sys, 'frozen', False):
        bundled = Path(sys._MEIPASS) / "ffmpeg.exe"
        if bundled.exists():
            _ffmpeg_path = str(bundled)
            return _ffmpeg_path

    # AppData (auto-downloaded)
    local = FFMPEG_DIR / "ffmpeg.exe"
    if local.exists():
        _ffmpeg_path = str(local)
        return _ffmpeg_path

    # System PATH
    which = shutil.which("ffmpeg")
    if which:
        _ffmpeg_path = which
        return _ffmpeg_path

    return None


def get_ffmpeg_dir():
    """Get directory containing ffmpeg, for yt-dlp's ffmpeg_location option"""
    path = get_ffmpeg_path()
    if path:
        return str(Path(path).parent)
    return None


def download_ffmpeg(log_callback=None, stop_check=None):
    """Download and install ffmpeg to AppData"""
    zip_path = FFMPEG_DIR / "ffmpeg_download.zip"

    try:
        FFMPEG_DIR.mkdir(parents=True, exist_ok=True)

        if log_callback:
            log_callback("Downloading ffmpeg (one-time download, ~130 MB)...")

        req = Request(FFMPEG_URL, headers={'User-Agent': 'BeamNG-PlaylistCreator/2.0'})
        response = urlopen(req, timeout=60)
        total = int(response.headers.get('Content-Length', 0))
        downloaded = 0
        chunk_size = 256 * 1024

        with open(zip_path, 'wb') as f:
            while True:
                if stop_check and stop_check():
                    zip_path.unlink(missing_ok=True)
                    return False
                chunk = response.read(chunk_size)
                if not chunk:
                    break
                f.write(chunk)
                downloaded += len(chunk)
                if log_callback and total > 0:
                    mb_done = downloaded / (1024 * 1024)
                    mb_total = total / (1024 * 1024)
                    pct = int((downloaded / total) * 100)
                    log_callback(f"Downloading ffmpeg... {mb_done:.0f}/{mb_total:.0f} MB ({pct}%)")

        if log_callback:
            log_callback("Extracting ffmpeg...")

        with zipfile.ZipFile(zip_path, 'r') as zf:
            for member in zf.namelist():
                fname = Path(member).name.lower()
                if fname in ('ffmpeg.exe', 'ffprobe.exe'):
                    target = FFMPEG_DIR / fname
                    with zf.open(member) as src, open(target, 'wb') as dst:
                        shutil.copyfileobj(src, dst)

        zip_path.unlink(missing_ok=True)

        ffmpeg_exe = FFMPEG_DIR / "ffmpeg.exe"
        if ffmpeg_exe.exists():
            global _ffmpeg_path
            _ffmpeg_path = str(ffmpeg_exe)
            result = subprocess.run(
                [str(ffmpeg_exe), "-version"], capture_output=True, timeout=5,
                startupinfo=_hidden_startupinfo()
            )
            if result.returncode == 0:
                if log_callback:
                    log_callback("ffmpeg installed successfully")
                return True

        if log_callback:
            log_callback("ERROR: ffmpeg extraction failed")
        return False

    except Exception as e:
        zip_path.unlink(missing_ok=True)
        if log_callback:
            log_callback(f"ERROR: Failed to install ffmpeg: {str(e)[:120]}")
        return False


def ensure_ffmpeg(log_callback=None, stop_check=None):
    """Find ffmpeg or auto-download it. Returns True if available."""
    path = get_ffmpeg_path()
    if path:
        try:
            result = subprocess.run(
                [path, "-version"], capture_output=True, timeout=5,
                startupinfo=_hidden_startupinfo()
            )
            if result.returncode == 0:
                if log_callback:
                    log_callback("ffmpeg ready")
                return True
        except Exception:
            pass

    if log_callback:
        log_callback("ffmpeg not found - installing automatically...")

    return download_ffmpeg(log_callback, stop_check)


# ============================================================================
# AUDIO PROCESSING
# ============================================================================

def _ydl_opts():
    """Common yt-dlp options for reliability"""
    opts = {
        'quiet': True,
        'no_warnings': True,
        'no_check_certificate': True,
        'socket_timeout': 30,
        'retries': 5,
        'fragment_retries': 5,
        'extractor_retries': 3,
        'file_access_retries': 3,
        'retry_sleep_functions': {'extractor': lambda n: 2 ** n},
    }
    ffdir = get_ffmpeg_dir()
    if ffdir:
        opts['ffmpeg_location'] = ffdir
    return opts


def _is_permanent_error(error_str):
    """Check if a yt-dlp error is permanent (no point retrying)"""
    keywords = [
        'private video', 'video unavailable', 'been removed', 'been deleted',
        'is not available', 'content is not available', 'not available in your country',
        'join this channel', 'members-only', 'age-restricted', 'age verification',
        'sign in to confirm', 'login required', 'copyright',
    ]
    lower = error_str.lower()
    return any(k in lower for k in keywords)


def create_bass_track(input_path, output_path, log_callback=None):
    """Create bass-boosted track using ffmpeg"""
    ffmpeg = get_ffmpeg_path() or "ffmpeg"
    try:
        result = subprocess.run(
            [ffmpeg, "-i", str(input_path), "-af", "lowpass=f=10:p=1,bass=g=10", "-y", str(output_path)],
            capture_output=True, check=True, timeout=300, startupinfo=_hidden_startupinfo()
        )
        return True
    except subprocess.TimeoutExpired:
        if log_callback:
            log_callback(f"Bass track timed out for {input_path.stem}")
        return False
    except subprocess.CalledProcessError as e:
        if log_callback:
            stderr = e.stderr.decode(errors='replace')[:150] if e.stderr else "No details"
            log_callback(f"Bass track error: {stderr}")
        return False
    except Exception as e:
        if log_callback:
            log_callback(f"Bass track error: {str(e)[:80]}")
        return False


def generate_amplitude_data(input_path, output_json, log_callback=None):
    """Generate JSON amplitude data for visualizer"""
    import numpy as np
    ffmpeg = get_ffmpeg_path() or "ffmpeg"
    try:
        raw_audio = subprocess.check_output(
            [ffmpeg, "-i", str(input_path), "-f", "s16le", "-acodec", "pcm_s16le",
             "-ac", "1", "-ar", "44100", "-"],
            stderr=subprocess.DEVNULL, startupinfo=_hidden_startupinfo()
        )

        if len(raw_audio) == 0:
            if log_callback:
                log_callback(f"Amplitude error: no audio data for {input_path.stem}")
            return False

        samples = np.frombuffer(raw_audio, dtype=np.int16).astype(np.int32)
        sample_rate = 44100
        total_rows = len(samples)
        song_length = total_rows / sample_rate

        target_count = 4000
        skip_factor = max(1, total_rows // target_count)
        values = [round(abs(samples[i]) / 32768.0, 2) for i in range(0, total_rows, skip_factor)]

        interval_beamng = (song_length / len(values)) * 1000

        with open(output_json, "w") as f:
            json.dump({"interval": round(interval_beamng, 9), "values": values}, f, separators=(',', ':'))

        return True
    except Exception as e:
        if log_callback:
            log_callback(f"Amplitude error: {str(e)[:80]}")
        return False


def get_audio_duration(file_path):
    """Get duration of audio file in seconds using mutagen"""
    try:
        ext = Path(file_path).suffix.lower()
        if ext == '.mp3':
            from mutagen.mp3 import MP3
            return int(MP3(file_path).info.length)
        elif ext == '.ogg':
            from mutagen.oggvorbis import OggVorbis
            return int(OggVorbis(file_path).info.length)
        elif ext == '.wav':
            from mutagen.wave import WAVE
            return int(WAVE(file_path).info.length)
        else:
            from mutagen import File
            audio = File(file_path)
            if audio and audio.info:
                return int(audio.info.length)
    except Exception:
        pass
    return 180  # fallback


# ============================================================================
# RADIOPACK BUILDING
# ============================================================================

def create_playlist_cover_from_pil(cover_images, output_path):
    """Create playlist cover from PIL Images — single cover if <4 songs, 2x2 grid if 4+"""
    valid = [img for img in cover_images if img is not None]
    if not valid:
        return False
    try:
        if len(valid) < 4:
            # Use one random cover scaled to fill
            cover = random.choice(valid)
            cover.resize((500, 500), Image.Resampling.LANCZOS).save(output_path)
        else:
            selected = random.sample(valid, 4)
            imgs = [img.resize((250, 250), Image.Resampling.LANCZOS) for img in selected]
            grid = Image.new('RGB', (500, 500))
            for i, img in enumerate(imgs):
                grid.paste(img, ((i % 2) * 250, (i // 2) * 250))
            grid.save(output_path)
        return True
    except Exception:
        return False


def generate_lua_playlist(music_dir, songs_dir, album_covers_dir, playlist_name,
                          songs_metadata, audio_format, log_callback=None):
    """Generate Lua playlist file for BeamNG"""
    audio_files = [f for f in music_dir.iterdir()
                   if f.suffix.lower() == f'.{audio_format}' and not f.stem.endswith('_bass')]
    if not audio_files:
        if log_callback:
            log_callback(f"ERROR: No {audio_format} files found after download")
        return False

    lua_var = playlist_name.lower().replace(' ', '_').replace('-', '_')
    display_name = playlist_name.replace('_', ' ')

    # Build metadata lookup from in-memory song data
    metadata = {}
    for s in songs_metadata:
        clean = clean_filename(s['title'])
        metadata[clean] = {'title': s['title'], 'artist': s['artist']}

    song_entries = []
    for audio_file in audio_files:
        name = audio_file.stem
        duration = get_audio_duration(audio_file)

        if name in metadata:
            title, artist = metadata[name]['title'], metadata[name]['artist']
        else:
            title = name.replace('_', ' ')
            artist = "Unknown"

        title = title.replace('"', '\\"').replace('\\', '\\\\')
        artist = artist.replace('"', '\\"').replace('\\', '\\\\')

        art_file = album_covers_dir / f"{name}_art.png"
        art_path = (f"local://local/vehicles/common/album_covers/{name}_art.png"
                     if art_file.exists() else
                     f"local://local/vehicles/common/album_covers/{lua_var}_playlist.png")

        song_entries.append(f"""    {{
        title = "{title}",
        artist = "{artist}",
        songPath = "art/sound/music/{name}",
		songBassPath = "art/sound/music/{name}_bass",
		songDataPath = "art/sound/music/{name}_data.json",
        albumArt = "{art_path}",
        duration = {duration}
    }}""")

    songs_text = ',\n'.join(song_entries)
    lua_content = f"""--Made with RedlineAuto & RoyalRenderings Auto Playlist Creator
-- https://www.patreon.com/SpoolingDieselDesigns
-- https://www.patreon.com/RedlineAuto

local playlistConfig = require("vehicles/common/lua/sdd_carplay_playlist_config")

local playlistName = "{display_name}"
local playlistCover = "local://local/vehicles/common/album_covers/{lua_var}_playlist.png"

local newSongs = {{
{songs_text}
}}

local defaultPlaylist = playlistConfig.playlists.default
local currentSize = defaultPlaylist and defaultPlaylist.songs and #defaultPlaylist.songs or 0
print("{lua_var} playlist file loaded, current playlist size: " .. currentSize)
playlistConfig.mergePlaylist("{lua_var}", playlistName, playlistCover, newSongs)
local {lua_var}Playlist = playlistConfig.playlists.{lua_var}
local newSize = {lua_var}Playlist and {lua_var}Playlist.songs and #{lua_var}Playlist.songs or 0
print("After merging {lua_var} playlist, size: " .. newSize)

return true"""

    lua_file = songs_dir / f"{lua_var}_playlist.lua"
    try:
        lua_file.write_text(lua_content, encoding='utf-8')
        if log_callback:
            log_callback(f"Lua playlist created: {lua_file.name} ({len(song_entries)} songs)")
        return True
    except Exception as e:
        if log_callback:
            log_callback(f"ERROR writing Lua file: {str(e)[:100]}")
        return False


def create_zip_archive(base_path, playlist_name, log_callback=None):
    """Create ZIP archive of the radiopack"""
    if log_callback:
        log_callback("Creating ZIP archive...")

    final_name, zip_path = find_available_name(
        playlist_name, base_path.parent, suffix="_Radiopack.zip", start_version=2, log_callback=log_callback)

    try:
        total_files = sum(1 for f in base_path.rglob('*') if f.is_file())
        current = 0

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for fp in base_path.rglob('*'):
                if fp.is_file():
                    current += 1
                    zf.write(fp, fp.relative_to(base_path))
                    if log_callback and current % 5 == 0:
                        pct = int((current / total_files) * 100)
                        log_callback(f"Archiving... {pct}%")

        if log_callback:
            size_mb = zip_path.stat().st_size / (1024 * 1024)
            log_callback(f"Created: {zip_path.name} ({size_mb:.1f} MB)")
        return zip_path
    except Exception as e:
        if log_callback:
            log_callback(f"ERROR creating ZIP: {str(e)[:100]}")
        return None


# ============================================================================
# DEPENDENCY CHECK
# ============================================================================

def _pip_install(package, upgrade=False, log_callback=None):
    """Install or upgrade a pip package. Retries with --user if permission denied."""
    action = "Upgrading" if upgrade else "Installing"
    if log_callback:
        log_callback(f"  {action} {package}...")

    for attempt_user in (False, True):
        try:
            cmd = [_python_exe(), "-m", "pip", "install", package, "-q", "--disable-pip-version-check"]
            if upgrade:
                cmd.insert(4, "--upgrade")
            if attempt_user:
                cmd.append("--user")

            result = subprocess.run(
                cmd, capture_output=True, timeout=120,
                startupinfo=_hidden_startupinfo()
            )

            if result.returncode != 0:
                stderr = result.stderr.decode(errors='replace')
                if not attempt_user and ("access is denied" in stderr.lower()
                                         or "permission" in stderr.lower()):
                    continue
                if log_callback:
                    log_callback(f"  Failed to install {package}: {stderr[:120]}")
                return False

            suffix = " (--user)" if attempt_user else ""
            if log_callback:
                log_callback(f"  {action.replace('ing', 'ed')} {package}{suffix}")
            return True
        except subprocess.TimeoutExpired:
            if log_callback:
                log_callback(f"  {package} install timed out")
            return False
        except Exception as e:
            if not attempt_user and "permission" in str(e).lower():
                continue
            if log_callback:
                log_callback(f"  Failed to install {package}: {str(e)[:100]}")
            return False

    return False


def _get_package_version(module_name):
    """Get installed version of a package"""
    try:
        mod = importlib.import_module(module_name)
        for attr in ('__version__', 'version', 'VERSION'):
            v = getattr(mod, attr, None)
            if v:
                return str(v) if isinstance(v, str) else getattr(v, '__version__', str(v))
        # Try version submodule
        try:
            ver_mod = importlib.import_module(f"{module_name}.version")
            return getattr(ver_mod, '__version__', None)
        except Exception:
            pass
    except Exception:
        pass
    return None


def check_and_install_dependencies(log_callback=None):
    """Check for required packages and install if missing.
    Skipped when running as a frozen exe (all deps are bundled)."""
    if getattr(sys, 'frozen', False):
        if log_callback:
            log_callback("All packages bundled")
        return True

    packages = {
        "requests": "requests",
        "yt_dlp": "yt-dlp",
        "PIL": "Pillow",
        "mutagen": "mutagen",
        "numpy": "numpy",
    }

    missing = [pip_name for mod, pip_name in packages.items()
               if importlib.util.find_spec(mod) is None]

    if missing:
        if log_callback:
            log_callback(f"Installing missing packages: {', '.join(missing)}")
        for pkg in missing:
            _pip_install(pkg, log_callback=log_callback)
    else:
        if log_callback:
            log_callback("All packages installed")

    # Silently update yt-dlp
    try:
        subprocess.run(
            [_python_exe(), "-m", "pip", "install", "--upgrade", "yt-dlp", "-q",
             "--disable-pip-version-check"],
            capture_output=True, timeout=30, startupinfo=_hidden_startupinfo()
        )
    except Exception:
        pass

    return True


def _get_ffmpeg_version():
    """Get a readable ffmpeg version string"""
    ffpath = get_ffmpeg_path()
    if not ffpath:
        return "not installed"
    try:
        res = subprocess.run(
            [ffpath, "-version"], capture_output=True, text=True, timeout=5,
            startupinfo=_hidden_startupinfo()
        )
        if res.returncode != 0 or not res.stdout:
            return "installed (unknown version)"
        first_line = res.stdout.split('\n')[0]
        parts = first_line.split()
        for i, p in enumerate(parts):
            if p == 'version' and i + 1 < len(parts):
                ver_raw = parts[i + 1]
                # Nightly: "N-118191-g6e42b68ec7-..." → "nightly (N-118191)"
                if ver_raw.startswith('N'):
                    segments = ver_raw.split('-')
                    if len(segments) >= 2:
                        return f"nightly ({segments[0]}-{segments[1]})"
                    return "nightly"
                # Release: "7.1.1-something" → "7.1.1"
                return ver_raw.split('-')[0]
        return "installed"
    except Exception:
        return "installed"


def _log_versions(log_callback):
    """Print all dependency versions to log"""
    log_callback("")
    log_callback("Current versions:")
    log_callback(f"  yt-dlp:        {_get_package_version('yt_dlp') or 'unknown'}")
    log_callback(f"  requests:      {_get_package_version('requests') or 'unknown'}")
    log_callback(f"  Pillow:        {_get_package_version('PIL') or 'unknown'}")
    log_callback(f"  mutagen:       {_get_package_version('mutagen') or 'unknown'}")
    log_callback(f"  numpy:         {_get_package_version('numpy') or 'unknown'}")
    log_callback(f"  customtkinter: {_get_package_version('customtkinter') or 'unknown'}")
    log_callback(f"  ffmpeg:        {_get_ffmpeg_version()}")


def update_packages(log_callback=None):
    """Check for outdated packages and upgrade only those that need it"""
    if log_callback:
        log_callback("=" * 50)
        log_callback("  Checking for Package Updates")
        log_callback("=" * 50)

    packages = ["yt-dlp", "requests", "Pillow", "mutagen", "numpy", "customtkinter"]

    # Find outdated packages
    if log_callback:
        log_callback("Checking for outdated packages...")

    outdated = []
    try:
        result = subprocess.run(
            [_python_exe(), "-m", "pip", "list", "--outdated", "--format=json",
             "--disable-pip-version-check"],
            capture_output=True, text=True, timeout=60,
            startupinfo=_hidden_startupinfo()
        )
        if result.returncode == 0 and result.stdout.strip():
            outdated_list = json.loads(result.stdout)
            outdated_names = {item['name'].lower() for item in outdated_list}
            outdated = [pkg for pkg in packages if pkg.lower() in outdated_names]

            # Log what we found
            for item in outdated_list:
                if item['name'].lower() in {p.lower() for p in packages}:
                    if log_callback:
                        log_callback(f"  {item['name']}: {item['version']} -> {item['latest_version']}")
    except Exception as e:
        if log_callback:
            log_callback(f"  Could not check outdated list: {str(e)[:80]}")
            log_callback("  Upgrading all packages instead...")
        outdated = packages

    if not outdated:
        if log_callback:
            log_callback("All packages are up to date!")
            _log_versions(log_callback)
        return True

    if log_callback:
        log_callback(f"Updating {len(outdated)} package(s)...")

    success = True
    for pkg in outdated:
        if not _pip_install(pkg, upgrade=True, log_callback=log_callback):
            success = False

    if log_callback:
        _log_versions(log_callback)
        log_callback("")
        if success:
            log_callback("All packages updated!")
        else:
            log_callback("Some packages failed - check messages above")

    return success


def update_ffmpeg(force=False, log_callback=None):
    """Update ffmpeg — skips if already installed unless force=True"""
    if log_callback:
        log_callback("=" * 50)
        log_callback("  Updating ffmpeg")
        log_callback("=" * 50)

    current_ver = _get_ffmpeg_version()

    if not force and get_ffmpeg_path():
        if log_callback:
            log_callback(f"ffmpeg is up to date: {current_ver}")
        return True

    if log_callback:
        if current_ver != "not installed":
            log_callback(f"Current version: {current_ver}")
        log_callback("Downloading latest ffmpeg...")

    # Remove old install
    if FFMPEG_DIR.exists():
        try:
            shutil.rmtree(FFMPEG_DIR)
        except Exception:
            pass

    global _ffmpeg_path
    _ffmpeg_path = None

    success = download_ffmpeg(log_callback)

    if log_callback and success:
        new_ver = _get_ffmpeg_version()
        log_callback(f"ffmpeg version: {new_ver}")

    return success


# ============================================================================
# PLAYLIST FETCHING  (single flat extract + parallel thumbnail downloads)
# ============================================================================

def fetch_playlist_info(playlist_url, progress_callback=None, stop_check=None):
    """
    Fast playlist load:
      1. One yt-dlp flat extract → all entries instantly
      2. Parse title / artist / duration from flat metadata
      3. Push every song to UI immediately (no covers yet)
      4. Download thumbnails in parallel, update covers as they arrive
    Returns (playlist_title, list_of_song_dicts)
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed
    import requests
    import yt_dlp

    playlist_title = "Unknown Playlist"

    # ── Single flat extract ──────────────────────────────────
    flat_opts = {
        'quiet': True,
        'no_warnings': True,
        'extract_flat': 'in_playlist',
        'no_check_certificate': True,
    }

    if progress_callback:
        progress_callback("status", "Fetching playlist...")

    with yt_dlp.YoutubeDL(flat_opts) as ydl:
        playlist_info = ydl.extract_info(playlist_url, download=False)

    if 'entries' not in playlist_info:
        raise ValueError("Could not find playlist entries")

    playlist_title = playlist_info.get('title', 'Unknown Playlist')
    entries = list(playlist_info['entries'])
    total = len(entries)

    if progress_callback:
        progress_callback("status", f"Found {total} songs in: {playlist_title}")
        progress_callback("total", total)

    # ── Build song list from flat data (instant) ─────────────
    songs = []
    for idx, entry in enumerate(entries):
        if stop_check and stop_check():
            break

        video_id = entry.get('id') or entry.get('url', '')

        # Parse title / artist from flat metadata
        raw_title = entry.get('title', f'Song {idx + 1}')
        artist = 'Unknown'

        # yt-dlp flat data may include these fields directly
        if entry.get('artist'):
            artist = entry['artist']
            title = entry.get('track') or raw_title
        elif entry.get('uploader'):
            artist = entry['uploader']
            title = raw_title
        elif ' - ' in raw_title:
            # YouTube Music often formats as "Artist - Title"
            parts = raw_title.split(' - ', 1)
            artist = parts[0].strip()
            title = parts[1].strip()
        else:
            title = raw_title

        # Clean "Topic" suffix from auto-generated YT Music channels
        if artist.endswith(' - Topic'):
            artist = artist[:-8].strip()

        # Duration from flat data
        dur_str = ''
        dur_sec = entry.get('duration')
        if dur_sec:
            try:
                mins, secs = divmod(int(float(dur_sec)), 60)
                dur_str = f"{mins}:{secs:02d}"
            except (ValueError, TypeError):
                pass

        song = {
            'title': title,
            'artist': artist,
            'duration': dur_str,
            'cover': None,
            'video_id': video_id,
            '_thumbnails': entry.get('thumbnails', []),  # Square art URLs from YT Music
        }
        songs.append(song)

        # Push to UI immediately (no cover yet)
        if progress_callback:
            progress_callback("song", (idx, song))

    # ── Download thumbnails in parallel ──────────────────────
    if progress_callback:
        progress_callback("status", f"Loading album art for {len(songs)} songs...")

    def _download_thumb(song_idx, vid_id, thumb_list):
        """Download a single thumbnail, preferring square album art from YT Music"""
        if not vid_id and not thumb_list:
            return song_idx, None

        # ── Priority 1: Square album art from flat extract thumbnails ──
        # YouTube Music provides square art via lh3.googleusercontent.com
        # Sort by size descending, prefer square
        square_urls = []
        other_urls = []

        for t in thumb_list:
            url = t.get('url', '')
            w = t.get('width', 0)
            h = t.get('height', 0)
            if not url:
                continue
            if w and h and w == h:
                square_urls.append((w, url))
            elif url:
                other_urls.append((max(w, h) if w and h else 0, url))

        # Try square thumbnails first (largest → smallest)
        square_urls.sort(key=lambda x: x[0], reverse=True)
        for _, thumb_url in square_urls:
            try:
                resp = requests.get(thumb_url, timeout=6)
                if resp.status_code == 200 and len(resp.content) > 1000:
                    img = Image.open(BytesIO(resp.content)).convert("RGB")
                    return song_idx, img
            except Exception:
                continue

        # ── Priority 2: Non-square thumbnails from flat extract (crop to square) ──
        other_urls.sort(key=lambda x: x[0], reverse=True)
        for _, thumb_url in other_urls:
            try:
                resp = requests.get(thumb_url, timeout=6)
                if resp.status_code == 200 and len(resp.content) > 1000:
                    img = Image.open(BytesIO(resp.content)).convert("RGB")
                    w, h = img.size
                    side = min(w, h)
                    left = (w - side) // 2
                    top = (h - side) // 2
                    img = img.crop((left, top, left + side, top + side))
                    return song_idx, img
            except Exception:
                continue

        # ── Priority 3: YouTube CDN fallback (always 16:9, crop to square) ──
        if vid_id:
            for thumb_url in [
                f"https://i.ytimg.com/vi/{vid_id}/hqdefault.jpg",
                f"https://i.ytimg.com/vi/{vid_id}/sddefault.jpg",
            ]:
                try:
                    resp = requests.get(thumb_url, timeout=6)
                    if resp.status_code == 200 and len(resp.content) > 1000:
                        img = Image.open(BytesIO(resp.content)).convert("RGB")
                        w, h = img.size
                        side = min(w, h)
                        left = (w - side) // 2
                        top = (h - side) // 2
                        img = img.crop((left, top, left + side, top + side))
                        return song_idx, img
                except Exception:
                    continue

        return song_idx, None

    with ThreadPoolExecutor(max_workers=10) as pool:
        futures = {
            pool.submit(_download_thumb, i, s['video_id'], s.get('_thumbnails', [])): i
            for i, s in enumerate(songs)
        }
        done_count = 0
        for future in as_completed(futures):
            if stop_check and stop_check():
                pool.shutdown(wait=False, cancel_futures=True)
                break
            try:
                song_idx, img = future.result()
                if img:
                    songs[song_idx]['cover'] = img
                    if progress_callback:
                        progress_callback("cover", (song_idx, img))
            except Exception:
                pass
            done_count += 1
            if progress_callback and done_count % 5 == 0:
                progress_callback("status", f"Album art: {done_count}/{len(songs)}")

    return playlist_title, songs


def fetch_single_song(song_url):
    """
    Fetch a single song's metadata + thumbnail.
    Returns a song dict: { title, artist, duration, cover, video_id }
    """
    import requests
    import yt_dlp

    opts = {
        'quiet': True,
        'no_warnings': True,
        'no_check_certificate': True,
    }

    with yt_dlp.YoutubeDL(opts) as ydl:
        info = ydl.extract_info(song_url, download=False)

    video_id = info.get('id', '')
    artist = (
        info.get('artist') or info.get('creator') or
        info.get('uploader', 'Unknown')
    )
    if artist.endswith(' - Topic'):
        artist = artist[:-8].strip()

    title = info.get('track') or info.get('title', 'Unknown')

    dur_str = ''
    dur_sec = info.get('duration')
    if dur_sec:
        mins, secs = divmod(int(dur_sec), 60)
        dur_str = f"{mins}:{secs:02d}"

    # Get thumbnail — prefer square from thumbnails list
    cover = None
    thumbnails = info.get('thumbnails', [])

    square_urls = []
    other_urls = []
    for t in thumbnails:
        url = t.get('url', '')
        w = t.get('width', 0)
        h = t.get('height', 0)
        if not url:
            continue
        if w and h and w == h:
            square_urls.append((w, url))
        elif url:
            other_urls.append((max(w, h) if w and h else 0, url))

    square_urls.sort(key=lambda x: x[0], reverse=True)
    other_urls.sort(key=lambda x: x[0], reverse=True)

    for url_list, crop in [(square_urls, False), (other_urls, True)]:
        for _, thumb_url in url_list:
            try:
                resp = requests.get(thumb_url, timeout=8)
                if resp.status_code == 200 and len(resp.content) > 1000:
                    img = Image.open(BytesIO(resp.content)).convert("RGB")
                    if crop:
                        w, h = img.size
                        side = min(w, h)
                        left = (w - side) // 2
                        top = (h - side) // 2
                        img = img.crop((left, top, left + side, top + side))
                    cover = img
                    break
            except Exception:
                continue
        if cover:
            break

    # CDN fallback
    if not cover and video_id:
        for thumb_url in [
            f"https://i.ytimg.com/vi/{video_id}/hqdefault.jpg",
            f"https://i.ytimg.com/vi/{video_id}/sddefault.jpg",
        ]:
            try:
                resp = requests.get(thumb_url, timeout=6)
                if resp.status_code == 200 and len(resp.content) > 1000:
                    img = Image.open(BytesIO(resp.content)).convert("RGB")
                    w, h = img.size
                    side = min(w, h)
                    left = (w - side) // 2
                    top = (h - side) // 2
                    cover = img.crop((left, top, left + side, top + side))
                    break
            except Exception:
                continue

    return {
        'title': title,
        'artist': artist,
        'duration': dur_str,
        'cover': cover,
        'video_id': video_id,
    }


# ============================================================================
# PLACEHOLDER COVER GENERATOR
# ============================================================================

# Cache the placeholder so we only generate it once
_placeholder_cache = {}

def make_placeholder_cover(size=120):
    """Generate a dark placeholder with a music note icon, matching YT Music style"""
    if size in _placeholder_cache:
        return _placeholder_cache[size].copy()

    bg_color = "#3a3a3a"
    note_color = "#8a8a8a"

    img = Image.new('RGB', (size, size), bg_color)
    draw = ImageDraw.Draw(img)

    cx = size // 2
    cy = size // 2

    # Music note proportions (scaled to image size)
    # Note head (oval/circle at bottom-left of stem)
    head_r = int(size * 0.10)
    head_cx = cx - int(size * 0.02)
    head_cy = cy + int(size * 0.14)

    # Stem (vertical line going up from note head)
    stem_x = head_cx + head_r - 1
    stem_top = cy - int(size * 0.18)
    stem_bottom = head_cy - int(head_r * 0.3)
    stem_width = max(2, int(size * 0.03))

    # Flag (small curve at top of stem) - simplified as angled line
    flag_len = int(size * 0.08)

    # Draw stem
    draw.rectangle(
        [stem_x - stem_width // 2, stem_top, stem_x + stem_width // 2, stem_bottom],
        fill=note_color
    )

    # Draw note head (filled ellipse, slightly wider than tall)
    draw.ellipse(
        [head_cx - int(head_r * 1.3), head_cy - head_r,
         head_cx + int(head_r * 1.3), head_cy + head_r],
        fill=note_color
    )

    # Draw a hollow center in the note head for the "d" look from the reference
    inner_r = int(head_r * 0.5)
    inner_cx = head_cx - int(head_r * 0.3)
    draw.ellipse(
        [inner_cx - int(inner_r * 1.2), head_cy - inner_r + 1,
         inner_cx + int(inner_r * 1.2), head_cy + inner_r - 1],
        fill=bg_color
    )

    _placeholder_cache[size] = img.copy()
    return img


def make_add_icon(size=120):
    """Generate a dark tile with a '+' icon for the Add Song button"""
    bg_color = "#2a2a2a"
    plus_color = "#606060"

    img = Image.new('RGB', (size, size), bg_color)
    draw = ImageDraw.Draw(img)

    cx, cy = size // 2, size // 2
    arm = int(size * 0.18)
    thickness = max(3, int(size * 0.045))

    # Horizontal bar
    draw.rectangle(
        [cx - arm, cy - thickness, cx + arm, cy + thickness],
        fill=plus_color
    )
    # Vertical bar
    draw.rectangle(
        [cx - thickness, cy - arm, cx + thickness, cy + arm],
        fill=plus_color
    )

    return img


# ============================================================================
# ADD SONG TILE (grid)
# ============================================================================

class AddSongTile(ctk.CTkFrame):
    """A tile in the grid that acts as an 'Add Song' button"""

    def __init__(self, parent, on_click, tile_size=128, **kwargs):
        super().__init__(parent, fg_color=("#e0e0e0", "#1e1e1e"), corner_radius=8, cursor="hand2", **kwargs)
        # No fixed width — stretches to fill grid cell
        self._on_click = on_click

        art_size = tile_size - 16
        img = make_add_icon(art_size)
        self._ctk_image = ctk.CTkImage(light_image=img, dark_image=img, size=(art_size, art_size))

        self._art_label = ctk.CTkLabel(self, image=self._ctk_image, text="")
        self._art_label.pack(padx=8, pady=(8, 4))

        self._title_label = ctk.CTkLabel(
            self, text="Add Song / Playlist",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("#555555", "#808080"), anchor="w"
        )
        self._title_label.pack(padx=8, anchor="w")

        self._sub_label = ctk.CTkLabel(
            self, text="Paste a link",
            font=ctk.CTkFont(size=10),
            text_color=("#888888", "#505050"), anchor="w"
        )
        self._sub_label.pack(padx=8, pady=(0, 6), anchor="w")

        # Click anywhere on the tile
        for w in [self, self._art_label, self._title_label, self._sub_label]:
            w.bind("<ButtonPress-1>", lambda e: self._on_click())
            w.bind("<Enter>", lambda e: self.configure(fg_color=("#d4d4d4", "#282828")))
            w.bind("<Leave>", lambda e: self.configure(fg_color=("#e0e0e0", "#1e1e1e")))


# ============================================================================
# ADD SONG ROW (list)
# ============================================================================

class AddSongRow(ctk.CTkFrame):
    """A row in the song list that acts as an 'Add Song' button"""

    def __init__(self, parent, on_click, **kwargs):
        super().__init__(parent, fg_color="transparent", height=36, cursor="hand2", **kwargs)
        self.grid_columnconfigure(2, weight=1)
        self.grid_propagate(False)
        self.configure(height=36)
        self._on_click = on_click

        # "+" icon
        self._plus_frame = ctk.CTkFrame(
            self, width=22, height=22, corner_radius=4,
            fg_color="transparent", border_width=2, border_color=("#cccccc", "#404040")
        )
        self._plus_frame.grid(row=0, column=0, padx=(8, 8), pady=7)
        self._plus_frame.grid_propagate(False)

        self._plus_label = ctk.CTkLabel(
            self._plus_frame, text="+",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=("#808080", "#606060"), anchor="center"
        )
        self._plus_label.place(relx=0.5, rely=0.5, anchor="center")

        # Text
        self._text_label = ctk.CTkLabel(
            self, text="Add Song / Playlist",
            font=ctk.CTkFont(size=12),
            text_color=("#808080", "#606060"), anchor="w"
        )
        self._text_label.grid(row=0, column=2, sticky="w", pady=4, padx=(6, 0))

        # Click + hover on all children
        for w in [self, self._plus_frame, self._plus_label, self._text_label]:
            w.bind("<ButtonPress-1>", lambda e: self._on_click())
            w.bind("<Enter>", self._on_enter, add="+")
            w.bind("<Leave>", self._on_leave, add="+")

    def _on_enter(self, event):
        self.configure(fg_color=("#e4e4e4", "#2a2a2a"))

    def _on_leave(self, event):
        try:
            mx, my = self.winfo_pointerx(), self.winfo_pointery()
            rx, ry = self.winfo_rootx(), self.winfo_rooty()
            rw, rh = self.winfo_width(), self.winfo_height()
            if not (rx <= mx <= rx + rw and ry <= my <= ry + rh):
                self.configure(fg_color="transparent")
        except Exception:
            self.configure(fg_color="transparent")


# ============================================================================
# ADD SONG / PLAYLIST POPUP
# ============================================================================

class AddSongDialog(ctk.CTkToplevel):
    """Popup dialog to paste a song or playlist URL. Stays open for multiple adds."""

    def __init__(self, parent, on_submit):
        super().__init__(parent)
        self.title("Add Song / Playlist")
        self.geometry("520x210")
        self.resizable(False, False)
        self.configure(fg_color=("#ffffff", "#111111"))
        self._on_submit = on_submit

        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width() - 520) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - 210) // 2
        self.geometry(f"+{px}+{py}")

        ctk.CTkLabel(
            self, text="Paste a YouTube Music link:",
            font=ctk.CTkFont(size=14, weight="bold"), text_color=("#1a1a1a", "#e0e0e0")
        ).pack(padx=20, pady=(20, 4), anchor="w")

        ctk.CTkLabel(
            self,
            text="Works with both single songs and full playlists. Duplicates are automatically skipped.",
            font=ctk.CTkFont(size=11), text_color=("#888888", "#505050"),
            wraplength=480, justify="left"
        ).pack(padx=20, pady=(0, 8), anchor="w")

        self.url_entry = ctk.CTkEntry(
            self, placeholder_text="https://music.youtube.com/watch?v=... or /playlist?list=...",
            height=38, font=ctk.CTkFont(size=12),
            border_width=0, corner_radius=6, fg_color=("#e8e8e8", "#1a1a1a")
        )
        self.url_entry.pack(fill="x", padx=20)

        self._status = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(size=11), text_color="#16a34a"
        )
        self._status.pack(padx=20, pady=(4, 0), anchor="w")

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(4, 20))

        ctk.CTkButton(
            btn_frame, text="Done", command=self.destroy,
            width=80, height=34, font=ctk.CTkFont(size=12),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color=("#d0d0d0", "#2e2e2e"), corner_radius=6,
            text_color=("#333333", "#d0d0d0")
        ).pack(side="right")

        ctk.CTkButton(
            btn_frame, text="Add", command=self._submit,
            width=100, height=34, font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#2563eb", hover_color="#1d4ed8", corner_radius=6
        ).pack(side="right", padx=(0, 8))

        # Auto-paste from clipboard
        try:
            clip = self.clipboard_get()
            if 'youtube' in clip or 'youtu.be' in clip or 'music.youtube' in clip:
                self.url_entry.insert(0, clip)
        except Exception:
            pass

        self.url_entry.focus_set()
        self.bind("<Return>", lambda e: self._submit())
        self.bind("<Escape>", lambda e: self.destroy())

    def _submit(self):
        url = self.url_entry.get().strip()
        if url:
            self._on_submit(url)
            self.url_entry.delete(0, "end")
            self._status.configure(text="Added! Paste another link or click Done.")
            self.url_entry.focus_set()


# ============================================================================
# TOOLTIP
# ============================================================================

class Tooltip:
    """Hover tooltip for any widget — shows a small floating label after a short delay"""

    def __init__(self, widget, text, delay=400):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._tip = None
        self._after_id = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None

    def _show(self):
        if self._tip:
            return
        x = self.widget.winfo_rootx() + self.widget.winfo_width() // 2
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._tip = tw = ctk.CTkToplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.configure(fg_color=("#e8e8e8", "#1a1a1a"))
        tw.attributes("-topmost", True)
        ctk.CTkLabel(
            tw, text=self.text, justify="left", anchor="w",
            font=ctk.CTkFont(size=11), text_color=("#404040", "#c0c0c0"),
            fg_color=("#e8e8e8", "#1a1a1a"), corner_radius=4,
            padx=8, pady=4
        ).pack(fill="x")

    def _hide(self, _event=None):
        self._cancel()
        if self._tip:
            self._tip.destroy()
            self._tip = None


# ============================================================================
# LOG WINDOW
# ============================================================================

class LogWindow(ctk.CTkToplevel):
    """Persistent log viewer — survives until app closes or user clears"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Log")
        self.geometry("700x420")
        self.resizable(True, True)
        self.minsize(400, 250)
        self.configure(fg_color=("#ffffff", "#111111"))
        self._parent = parent
        self.after(200, self._build_ui)

    def _build_ui(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=12, pady=12)
        container.grid_rowconfigure(1, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # ── Top bar: buttons ──
        btn_bar = ctk.CTkFrame(container, fg_color="transparent")
        btn_bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        ctk.CTkLabel(
            btn_bar, text="Log",
            font=ctk.CTkFont(size=14, weight="bold"), text_color=("#333333", "#d0d0d0")
        ).pack(side="left")

        ctk.CTkButton(
            btn_bar, text="Close", command=self._close_window,
            width=70, height=28, font=ctk.CTkFont(size=11),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color=("#d0d0d0", "#2e2e2e"),
            corner_radius=6, text_color=("#333333", "#d0d0d0")
        ).pack(side="right", padx=(4, 0))

        ctk.CTkButton(
            btn_bar, text="Clear", command=self._clear_log,
            width=70, height=28, font=ctk.CTkFont(size=11),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color=("#d0d0d0", "#2e2e2e"),
            corner_radius=6, text_color=("#333333", "#d0d0d0")
        ).pack(side="right", padx=(4, 0))

        ctk.CTkButton(
            btn_bar, text="Copy All", command=self._copy_all,
            width=80, height=28, font=ctk.CTkFont(size=11),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color=("#d0d0d0", "#2e2e2e"),
            corner_radius=6, text_color=("#333333", "#d0d0d0")
        ).pack(side="right", padx=(4, 0))

        # ── Text area (supports selection + scrolling) ──
        self._textbox = ctk.CTkTextbox(
            container, font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=("#f5f5f5", "#0d0d0d"), text_color=("#222222", "#cccccc"),
            border_width=1, border_color=("#d0d0d0", "#2a2a2a"),
            corner_radius=6, wrap="word", state="disabled"
        )
        self._textbox.grid(row=1, column=0, sticky="nsew")

        # Replay any buffered log lines
        if hasattr(self._parent, '_log_buffer'):
            self._textbox.configure(state="normal")
            self._textbox.insert("end", self._parent._log_buffer)
            self._textbox.see("end")
            self._textbox.configure(state="disabled")

        # Center + bring to front
        self.update_idletasks()
        x = self._parent.winfo_x() + (self._parent.winfo_width() - self.winfo_width()) // 2
        y = self._parent.winfo_y() + (self._parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
        self.attributes("-topmost", True)
        self.after(100, lambda: self.attributes("-topmost", False))

    def append(self, text):
        """Append a line to the log (thread-safe — caller must use self.after)"""
        if not hasattr(self, '_textbox'):
            return
        try:
            self._textbox.configure(state="normal")
            self._textbox.insert("end", text + "\n")
            self._textbox.see("end")
            self._textbox.configure(state="disabled")
        except Exception:
            pass

    def _copy_all(self):
        if not hasattr(self, '_textbox'):
            return
        self._textbox.configure(state="normal")
        content = self._textbox.get("1.0", "end").strip()
        self._textbox.configure(state="disabled")
        self.clipboard_clear()
        self.clipboard_append(content)

    def _clear_log(self):
        if not hasattr(self, '_textbox'):
            return
        self._textbox.configure(state="normal")
        self._textbox.delete("1.0", "end")
        self._textbox.configure(state="disabled")
        # Also clear the buffer
        if hasattr(self._parent, '_log_buffer'):
            self._parent._log_buffer = ""

    def _close_window(self):
        self.withdraw()


# ============================================================================
# SETTINGS DIALOG
# ============================================================================

class SettingsDialog(ctk.CTkToplevel):
    """Application settings — persisted to JSON"""

    def __init__(self, parent, settings, on_save):
        super().__init__(parent)
        self.title("Settings")
        self.geometry("520x540")
        self.resizable(False, False)
        self.configure(fg_color=("#ffffff", "#111111"))
        self._settings = dict(settings)  # work on a copy
        self._on_save = on_save
        self._parent = parent

        # Defer all widget creation to work around CTkToplevel rendering bug on Windows
        self.after(200, self._build_ui)

    def _build_ui(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=20, pady=(12, 20))
        self._container = container

        # ── Output Folder ────────────────────────────────────
        self._section_label("Output Folder")

        folder_frame = ctk.CTkFrame(container, fg_color="transparent")
        folder_frame.pack(fill="x", pady=(0, 4))
        folder_frame.grid_columnconfigure(0, weight=1)

        self._folder_entry = ctk.CTkEntry(
            folder_frame, height=34, font=ctk.CTkFont(size=11),
            border_width=0, corner_radius=6, fg_color=("#e8e8e8", "#1a1a1a"),
            state="normal"
        )
        self._folder_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._folder_entry.insert(0, self._settings["output_folder"])
        self._folder_entry.configure(state="disabled")

        ctk.CTkButton(
            folder_frame, text="Browse", command=self._browse_folder,
            width=70, height=34, font=ctk.CTkFont(size=11),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color=("#d0d0d0", "#2e2e2e"), corner_radius=6,
            text_color=("#333333", "#d0d0d0")
        ).grid(row=0, column=1, padx=(0, 4))

        ctk.CTkButton(
            folder_frame, text="Reset", command=self._reset_folder,
            width=58, height=34, font=ctk.CTkFont(size=11),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color=("#d0d0d0", "#2e2e2e"), corner_radius=6,
            text_color=("#333333", "#d0d0d0")
        ).grid(row=0, column=2)

        # ── Audio Quality ────────────────────────────────────
        self._section_label("Audio Quality")

        self._quality_var = ctk.StringVar(value=f"{self._settings['audio_quality']} kbps")
        quality_menu = ctk.CTkOptionMenu(
            container, values=["128 kbps", "192 kbps", "256 kbps", "320 kbps"],
            variable=self._quality_var,
            command=self._on_quality_change,
            font=ctk.CTkFont(size=12), width=140,
            fg_color=("#e8e8e8", "#1a1a1a"), button_color=("#d0d0d0", "#1e1e1e"),
            button_hover_color=("#d0d0d0", "#2e2e2e"), dropdown_fg_color=("#e8e8e8", "#1a1a1a"),
            text_color=("#333333", "#d0d0d0")
        )
        quality_menu.pack(anchor="w", pady=(0, 4))
        Tooltip(quality_menu, "Higher bitrate = better quality but larger files. 192 is a good balance.")

        # ── Appearance Mode ──────────────────────────────────
        self._section_label("Appearance")

        self._appearance_var = ctk.StringVar(value=self._settings["appearance"])
        # Wrap in frame — CTkSegmentedButton doesn't support .bind()
        appear_wrap = ctk.CTkFrame(container, fg_color="transparent")
        appear_wrap.pack(anchor="w", pady=(0, 8))
        ctk.CTkSegmentedButton(
            appear_wrap, values=["Dark", "Light", "System"],
            variable=self._appearance_var,
            command=self._on_appearance_change,
            font=ctk.CTkFont(size=12),
            fg_color=("#e8e8e8", "#1a1a1a"), selected_color="#2563eb",
            selected_hover_color="#1d4ed8", unselected_color=("#e0e0e0", "#1e1e1e"),
            unselected_hover_color=("#d0d0d0", "#2e2e2e"),
            text_color=("#333333", "#d0d0d0")
        ).pack()

        # ── Packages ─────────────────────────────────────────
        self._section_label("Packages")

        self._dep_row = ctk.CTkFrame(container, fg_color="transparent")
        self._dep_row.pack(fill="x", pady=(0, 4))

        self._dep_status = ctk.CTkLabel(
            self._dep_row, text="",
            font=ctk.CTkFont(size=11), text_color=("#888888", "#505050")
        )

        self._build_dep_buttons()

        # ── ffmpeg ───────────────────────────────────────────
        self._section_label("ffmpeg")

        self._ffmpeg_row = ctk.CTkFrame(container, fg_color="transparent")
        self._ffmpeg_row.pack(fill="x", pady=(0, 8))

        self._ffmpeg_status = ctk.CTkLabel(
            self._ffmpeg_row, text="",
            font=ctk.CTkFont(size=11), text_color=("#888888", "#505050")
        )

        self._build_ffmpeg_buttons()

        # ── App Update ──────────────────────────────────────
        self._section_label("App Update")

        self._update_row = ctk.CTkFrame(container, fg_color="transparent")
        self._update_row.pack(fill="x", pady=(0, 4))

        self._maint_btn(self._update_row, "Check for Updates",
                        self._check_app_update).pack(side="left")

        self._update_status = ctk.CTkLabel(
            self._update_row, text=f"v{__version__}",
            font=ctk.CTkFont(size=11), text_color=("#888888", "#505050")
        )
        self._update_status.pack(side="left", padx=(12, 0))

        # ── Close button ─────────────────────────────────────
        ctk.CTkButton(
            container, text="Close", command=self.destroy,
            width=80, height=34, font=ctk.CTkFont(size=12),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color=("#d0d0d0", "#2e2e2e"), corner_radius=6,
            text_color=("#333333", "#d0d0d0")
        ).pack(side="bottom", pady=(8, 0))

        self.bind("<Escape>", lambda e: self.destroy())

        # Center on parent and bring to front (no grab_set or transient — avoids white titlebar on Windows)
        self.update_idletasks()
        px = self._parent.winfo_rootx() + (self._parent.winfo_width() - 520) // 2
        py = self._parent.winfo_rooty() + (self._parent.winfo_height() - 500) // 2
        self.geometry(f"+{px}+{py}")
        # Flash topmost to force window to front on Windows, then disable so it's not always-on-top
        self.attributes("-topmost", True)
        self.after(100, lambda: self.attributes("-topmost", False))
        self.focus_force()

    def _section_label(self, text):
        lbl = ctk.CTkLabel(
            self._container, text=text,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=("#333333", "#aaaaaa"), anchor="w"
        )
        lbl.pack(pady=(8, 4), anchor="w")
        return lbl

    def _save(self):
        save_settings(self._settings)
        self._on_save(self._settings)

    def _set_folder_text(self, path):
        self._folder_entry.configure(state="normal")
        self._folder_entry.delete(0, "end")
        self._folder_entry.insert(0, path)
        self._folder_entry.configure(state="disabled")

    def _browse_folder(self):
        folder = filedialog.askdirectory(
            initialdir=self._settings["output_folder"],
            title="Select Output Folder"
        )
        if folder:
            self._settings["output_folder"] = folder
            self._set_folder_text(folder)
            self._save()

    def _reset_folder(self):
        default = _default_output_folder()
        self._settings["output_folder"] = default
        self._set_folder_text(default)
        self._save()

    def _on_quality_change(self, value):
        self._settings["audio_quality"] = value.replace(" kbps", "")
        self._save()

    def _on_appearance_change(self, value):
        self._settings["appearance"] = value
        self._save()

    def _maint_log(self, status_label, msg):
        """Route message to both a status label and the main log"""
        self.after(0, lambda m=msg: status_label.configure(text=m))
        if hasattr(self._parent, '_log'):
            self._parent._log(msg)

    def _maint_btn(self, parent, text, command, width=140):
        """Create a styled maintenance button"""
        return ctk.CTkButton(
            parent, text=text, command=command,
            width=width, height=32, font=ctk.CTkFont(size=11),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color=("#d0d0d0", "#2e2e2e"),
            corner_radius=6, text_color=("#333333", "#d0d0d0")
        )

    # ── Packages dynamic buttons ──

    def _check_deps_installed(self):
        """Check if all required packages are installed"""
        modules = ["requests", "yt_dlp", "PIL", "mutagen", "numpy"]
        return all(importlib.util.find_spec(m) is not None for m in modules)

    def _build_dep_buttons(self):
        """Rebuild package buttons based on install state"""
        for w in self._dep_row.winfo_children():
            w.destroy()

        if self._check_deps_installed():
            self._maint_btn(self._dep_row, "Check for Updates",
                            self._update_deps).pack(side="left")
            self._maint_btn(self._dep_row, "Uninstall",
                            self._uninstall_deps, width=100).pack(side="left", padx=(6, 0))
        else:
            self._maint_btn(self._dep_row, "Install Packages",
                            self._install_deps).pack(side="left")

        self._dep_status = ctk.CTkLabel(
            self._dep_row, text="",
            font=ctk.CTkFont(size=11), text_color=("#888888", "#505050")
        )
        self._dep_status.pack(side="left", padx=(12, 0))

    def _install_deps(self):
        self._dep_status.configure(text="Installing...", text_color=("#1a1a1a", "#e0e0e0"))

        def _run():
            try:
                check_and_install_dependencies(
                    log_callback=lambda msg: self._maint_log(self._dep_status, msg))
                self.after(0, lambda: self._dep_status.configure(
                    text="Installed!", text_color="#16a34a"))
                self.after(0, self._build_dep_buttons)
            except Exception as e:
                self.after(0, lambda: self._dep_status.configure(
                    text=f"Error: {str(e)[:60]}", text_color="#ef4444"))

        threading.Thread(target=_run, daemon=True).start()

    def _update_deps(self):
        self._dep_status.configure(text="Updating...", text_color=("#1a1a1a", "#e0e0e0"))

        def _run():
            try:
                success = update_packages(
                    log_callback=lambda msg: self._maint_log(self._dep_status, msg))
                if success:
                    self.after(0, lambda: self._dep_status.configure(
                        text="Up to date!", text_color="#16a34a"))
                else:
                    self.after(0, lambda: self._dep_status.configure(
                        text="Some failed", text_color="#ef4444"))
            except Exception as e:
                self.after(0, lambda: self._dep_status.configure(
                    text=f"Error: {str(e)[:60]}", text_color="#ef4444"))

        threading.Thread(target=_run, daemon=True).start()

    def _uninstall_deps(self):
        self._dep_status.configure(text="Uninstalling...", text_color=("#1a1a1a", "#e0e0e0"))

        def _run():
            packages = ["yt-dlp", "requests", "mutagen", "numpy"]
            # Don't uninstall Pillow or customtkinter — the app needs them to run
            log = lambda msg: self._maint_log(self._dep_status, msg)
            log("=" * 50)
            log("  Uninstalling Packages")
            log("=" * 50)
            for pkg in packages:
                try:
                    log(f"  Uninstalling {pkg}...")
                    result = subprocess.run(
                        [_python_exe(), "-m", "pip", "uninstall", pkg, "-y", "-q"],
                        capture_output=True, timeout=60,
                        startupinfo=_hidden_startupinfo()
                    )
                    if result.returncode == 0:
                        log(f"  Uninstalled {pkg}")
                    else:
                        log(f"  Failed to uninstall {pkg}")
                except Exception as e:
                    log(f"  Error: {str(e)[:80]}")

            log("")
            log("Note: Pillow and customtkinter kept (required by app)")
            log("Uninstall complete")
            self.after(0, lambda: self._dep_status.configure(
                text="Uninstalled", text_color="#16a34a"))
            self.after(0, self._build_dep_buttons)

        threading.Thread(target=_run, daemon=True).start()

    # ── ffmpeg dynamic buttons ──

    def _build_ffmpeg_buttons(self):
        """Rebuild ffmpeg buttons based on install state"""
        for w in self._ffmpeg_row.winfo_children():
            w.destroy()

        if get_ffmpeg_path():
            self._maint_btn(self._ffmpeg_row, "Check for Updates",
                            self._update_ffmpeg).pack(side="left")
            self._maint_btn(self._ffmpeg_row, "Uninstall",
                            self._uninstall_ffmpeg, width=100).pack(side="left", padx=(6, 0))
            ver = _get_ffmpeg_version()
            self._ffmpeg_status = ctk.CTkLabel(
                self._ffmpeg_row, text=ver,
                font=ctk.CTkFont(size=11), text_color=("#888888", "#505050")
            )
        else:
            self._maint_btn(self._ffmpeg_row, "Install ffmpeg",
                            self._install_ffmpeg).pack(side="left")
            self._ffmpeg_status = ctk.CTkLabel(
                self._ffmpeg_row, text="Not installed",
                font=ctk.CTkFont(size=11), text_color=("#888888", "#505050")
            )

        self._ffmpeg_status.pack(side="left", padx=(12, 0))

    def _install_ffmpeg(self):
        self._ffmpeg_status.configure(text="Installing...", text_color=("#1a1a1a", "#e0e0e0"))

        def _run():
            try:
                success = update_ffmpeg(
                    force=True,
                    log_callback=lambda msg: self._maint_log(self._ffmpeg_status, msg))
                if success:
                    self.after(0, lambda: self._ffmpeg_status.configure(
                        text="Installed!", text_color="#16a34a"))
                else:
                    self.after(0, lambda: self._ffmpeg_status.configure(
                        text="Failed", text_color="#ef4444"))
                self.after(200, self._build_ffmpeg_buttons)
            except Exception as e:
                self.after(0, lambda: self._ffmpeg_status.configure(
                    text=f"Error: {str(e)[:60]}", text_color="#ef4444"))

        threading.Thread(target=_run, daemon=True).start()

    def _update_ffmpeg(self):
        self._ffmpeg_status.configure(text="Checking...", text_color=("#1a1a1a", "#e0e0e0"))

        def _run():
            try:
                success = update_ffmpeg(
                    force=False,
                    log_callback=lambda msg: self._maint_log(self._ffmpeg_status, msg))
                if success:
                    self.after(0, lambda: self._ffmpeg_status.configure(
                        text="Up to date!", text_color="#16a34a"))
                else:
                    self.after(0, lambda: self._ffmpeg_status.configure(
                        text="Failed", text_color="#ef4444"))
                self.after(200, self._build_ffmpeg_buttons)
            except Exception as e:
                self.after(0, lambda: self._ffmpeg_status.configure(
                    text=f"Error: {str(e)[:60]}", text_color="#ef4444"))

        threading.Thread(target=_run, daemon=True).start()

    def _uninstall_ffmpeg(self):
        self._ffmpeg_status.configure(text="Uninstalling...", text_color=("#1a1a1a", "#e0e0e0"))

        def _run():
            log = lambda msg: self._maint_log(self._ffmpeg_status, msg)
            log("Uninstalling ffmpeg...")
            global _ffmpeg_path
            if FFMPEG_DIR.exists():
                try:
                    shutil.rmtree(FFMPEG_DIR)
                    log("Removed ffmpeg from AppData")
                except Exception as e:
                    log(f"Error removing ffmpeg: {str(e)[:80]}")
            _ffmpeg_path = None
            log("ffmpeg uninstalled")
            self.after(0, lambda: self._ffmpeg_status.configure(
                text="Uninstalled", text_color="#16a34a"))
            self.after(200, self._build_ffmpeg_buttons)

        threading.Thread(target=_run, daemon=True).start()

    # ── App Update ──
    def _check_app_update(self):
        self._update_status.configure(text="Checking...", text_color=("#1a1a1a", "#e0e0e0"))

        def _run():
            try:
                info = check_for_update(
                    log_callback=lambda msg: self._maint_log(self._update_status, msg))
                if info is None:
                    self.after(0, lambda: self._update_status.configure(
                        text=f"v{__version__} — up to date!", text_color="#16a34a"))
                else:
                    self.after(0, lambda: self._offer_update(info))
            except Exception as e:
                self.after(0, lambda: self._update_status.configure(
                    text=f"Error: {str(e)[:60]}", text_color="#ef4444"))

        threading.Thread(target=_run, daemon=True).start()

    def _offer_update(self, update_info):
        result = messagebox.askyesno(
            "Update Available",
            f"A new version ({update_info['version']}) is available.\n\n"
            f"You are running v{__version__}.\n\n"
            f"Would you like to update now?",
            parent=self
        )
        if result:
            self._update_status.configure(text="Downloading...", text_color=("#1a1a1a", "#e0e0e0"))

            def _dl():
                def _dl_log(msg):
                    self._maint_log(self._update_status, msg)

                path = download_update(
                    update_info["download_url"],
                    update_info["asset_name"],
                    log_callback=_dl_log
                )
                if path:
                    subprocess.Popen(
                        [str(path), '/SILENT', '/CLOSEAPPLICATIONS'],
                        startupinfo=_hidden_startupinfo()
                    )
                    self.after(500, lambda: self._parent.destroy())
                else:
                    self.after(0, lambda: self._update_status.configure(
                        text="Download failed", text_color="#ef4444"))

            threading.Thread(target=_dl, daemon=True).start()


# ============================================================================
# SONG GRID TILE
# ============================================================================

class SongTile(ctk.CTkFrame):
    """A single album art tile in the grid — expands to fill its cell"""

    def __init__(self, parent, song_data, tile_size=128, **kwargs):
        super().__init__(parent, fg_color=("#e0e0e0", "#1e1e1e"), corner_radius=8, **kwargs)
        # No fixed width — tile stretches to fill grid cell

        art_size = tile_size - 16

        if song_data.get("cover") and isinstance(song_data["cover"], Image.Image):
            img = song_data["cover"].resize((art_size, art_size), Image.Resampling.LANCZOS)
        else:
            img = make_placeholder_cover(art_size)

        self._ctk_image = ctk.CTkImage(light_image=img, dark_image=img, size=(art_size, art_size))

        art_label = ctk.CTkLabel(self, image=self._ctk_image, text="")
        art_label.pack(pady=(8, 4))

        title_label = ctk.CTkLabel(
            self, text=song_data["title"],
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("#1a1a1a", "#e0e0e0"),
            anchor="w"
        )
        title_label.pack(padx=8, anchor="w")

        artist_label = ctk.CTkLabel(
            self, text=song_data["artist"],
            font=ctk.CTkFont(size=10),
            text_color=("#555555", "#808080"),
            anchor="w"
        )
        artist_label.pack(padx=8, pady=(0, 6), anchor="w")


# ============================================================================
# SONG LIST ROW
# ============================================================================

class SongRow(ctk.CTkFrame):
    """
    A single row in the song list.
    The ENTIRE row is one click/drag target. No interactive checkbox widget —
    just a visual check indicator that the drag manager controls.
    Hover shows an X button for individual removal.
    """

    _currently_hovered = None  # class-level: ensures only one row is highlighted at a time

    def __init__(self, parent, index, song_data, on_toggle=None, on_remove=None, **kwargs):
        super().__init__(parent, fg_color="transparent", height=36, **kwargs)
        self.grid_columnconfigure(2, weight=1)
        self.grid_propagate(False)
        self.configure(height=36)

        self.selected = ctk.BooleanVar(value=True)
        self.on_toggle = on_toggle
        self.on_remove = on_remove
        self.index = index

        self._default_color = "transparent"
        self._hover_color = ("#e4e4e4", "#2a2a2a")

        # Visual check indicator
        self._check_frame = ctk.CTkFrame(
            self, width=22, height=22, corner_radius=4,
            fg_color="#16a34a", border_width=2, border_color="#16a34a"
        )
        self._check_frame.grid(row=0, column=0, padx=(8, 8), pady=7)
        self._check_frame.grid_propagate(False)

        self._check_label = ctk.CTkLabel(
            self._check_frame, text="\u2713",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("#111111", "#ffffff"), anchor="center"
        )
        self._check_label.place(relx=0.5, rely=0.5, anchor="center")

        # Index number
        self._idx_label = ctk.CTkLabel(
            self, text=f"{index}.",
            font=ctk.CTkFont(size=11),
            text_color=("#888888", "#505050"), width=28, anchor="e"
        )
        self._idx_label.grid(row=0, column=1, padx=(0, 6), sticky="e")

        # Title + Artist
        self._info_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._info_frame.grid(row=0, column=2, sticky="ew", pady=4)
        self._info_frame.grid_columnconfigure(0, weight=1)

        text = f"{song_data['title']}  --  {song_data['artist']}"
        self.info_label = ctk.CTkLabel(
            self._info_frame, text=text,
            font=ctk.CTkFont(size=12),
            text_color=("#2a2a2a", "#d0d0d0"), anchor="w"
        )
        self.info_label.grid(row=0, column=0, sticky="w")

        # Duration
        self._dur_label = ctk.CTkLabel(
            self, text=song_data.get("duration", ""),
            font=ctk.CTkFont(size=11, family="Consolas"),
            text_color=("#808080", "#606060"), width=45, anchor="e"
        )
        self._dur_label.grid(row=0, column=3, padx=(8, 0), sticky="e")

        # Remove button (hidden by default, shows on hover)
        self._remove_btn = ctk.CTkLabel(
            self, text="\u00d7", width=24,
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=("#aaaaaa", "#404040"), anchor="center", cursor="hand2"
        )
        self._remove_btn.grid(row=0, column=4, padx=(2, 8), sticky="e")
        self._remove_btn.grid_remove()  # Hidden initially
        self._remove_btn.bind("<ButtonPress-1>", self._on_remove_click)
        self._remove_btn.bind("<Enter>", lambda e: self._remove_btn.configure(text_color="#ef4444"))
        self._remove_btn.bind("<Leave>", lambda e: self._remove_btn.configure(text_color=("#aaaaaa", "#404040")))

        # Bind hover on all children
        for w in self._get_all_widgets():
            w.bind("<Enter>", self._on_enter, add="+")
            w.bind("<Leave>", self._on_leave, add="+")

        self._update_check_visual()

    def _get_all_widgets(self):
        return [self, self._check_frame, self._check_label,
                self._idx_label, self._info_frame, self.info_label,
                self._dur_label, self._remove_btn]

    def _update_check_visual(self):
        if self.selected.get():
            self._check_frame.configure(fg_color="#16a34a", border_color="#16a34a")
            self._check_label.configure(text="\u2713", text_color=("#111111", "#ffffff"))
        else:
            self._check_frame.configure(fg_color="transparent", border_color=("#cccccc", "#404040"))
            self._check_label.configure(text="", text_color=("#aaaaaa", "#404040"))

    def _on_enter(self, event):
        # Un-highlight whichever row was previously hovered
        prev = SongRow._currently_hovered
        if prev is not None and prev is not self:
            try:
                prev.configure(fg_color=prev._default_color)
                prev._remove_btn.grid_remove()
            except Exception:
                pass
        SongRow._currently_hovered = self
        self.configure(fg_color=self._hover_color)
        self._remove_btn.grid()

    def _on_leave(self, event):
        try:
            mx, my = self.winfo_pointerx(), self.winfo_pointery()
            rx, ry = self.winfo_rootx(), self.winfo_rooty()
            rw, rh = self.winfo_width(), self.winfo_height()
            if not (rx <= mx <= rx + rw and ry <= my <= ry + rh):
                self.configure(fg_color=self._default_color)
                self._remove_btn.grid_remove()
                if SongRow._currently_hovered is self:
                    SongRow._currently_hovered = None
        except Exception:
            self.configure(fg_color=self._default_color)
            self._remove_btn.grid_remove()
            if SongRow._currently_hovered is self:
                SongRow._currently_hovered = None

    def _on_remove_click(self, event):
        if self.on_remove:
            self.on_remove(self.index)
        return "break"  # Don't trigger drag-select

    def set_selected(self, value: bool):
        self.selected.set(value)
        self._update_check_visual()

    def toggle(self):
        new_val = not self.selected.get()
        self.set_selected(new_val)
        if self.on_toggle:
            self.on_toggle()
        return new_val

    def update_index(self, new_index):
        self.index = new_index
        self._idx_label.configure(text=f"{new_index}.")


# ============================================================================
# DRAG-SELECT MANAGER
# ============================================================================

class DragSelectManager:
    """
    Click-and-drag multi-select on the song list.
    Click ANYWHERE on a row (checkbox, text, duration, etc.) to toggle it,
    then drag up/down to apply that same state to other rows.
    """

    def __init__(self, scroll_frame, get_rows_fn, on_change_fn):
        self.scroll_frame = scroll_frame
        self.get_rows = get_rows_fn
        self.on_change = on_change_fn

        self._dragging = False
        self._drag_state = True   # True = check, False = uncheck
        self._last_hit_index = -1

        root = scroll_frame.winfo_toplevel()
        root.bind("<ButtonPress-1>", self._on_press, add="+")
        root.bind("<B1-Motion>", self._on_motion, add="+")
        root.bind("<ButtonRelease-1>", self._on_release, add="+")

    def _find_parent_row(self, widget):
        """Walk up the widget tree to find the SongRow ancestor"""
        w = widget
        while w:
            if isinstance(w, SongRow):
                return w
            try:
                w = w.master
            except Exception:
                break
        return None

    def _get_row_at_y(self, abs_y):
        """Find which SongRow the cursor is over by screen y"""
        for row in self.get_rows():
            try:
                ry = row.winfo_rooty()
                rh = row.winfo_height()
                if ry <= abs_y <= ry + rh:
                    return row
            except Exception:
                continue
        return None

    def _on_press(self, event):
        # Don't start drag if clicking the remove button
        w = event.widget
        if hasattr(w, '_name') and 'remove' in str(getattr(w, '_name', '')):
            return
        # Check if widget is a SongRow's _remove_btn
        try:
            if w.master and hasattr(w.master, '_remove_btn') and w is w.master._remove_btn:
                return
        except Exception:
            pass
        # Also check by walking up and looking for on_remove attribute
        parent_row = self._find_parent_row(event.widget)
        if parent_row and hasattr(parent_row, '_remove_btn'):
            # Check if click is in the remove button's area
            try:
                btn = parent_row._remove_btn
                bx = btn.winfo_rootx()
                bw = btn.winfo_width()
                mx = event.x_root
                if bx <= mx <= bx + bw and btn.winfo_ismapped():
                    return  # Click was on remove button area
            except Exception:
                pass

        row = parent_row
        if not row:
            self._dragging = False
            return

        # Toggle the clicked row, start drag in that direction
        new_state = row.toggle()
        self._dragging = True
        self._drag_state = new_state
        self._last_hit_index = row.index
        self.on_change()

        # Prevent the event from reaching the checkbox's own handler
        return "break"

    def _on_motion(self, event):
        if not self._dragging:
            return

        row = self._get_row_at_y(event.y_root)
        if row and row.index != self._last_hit_index:
            rows = self.get_rows()
            lo = min(self._last_hit_index, row.index)
            hi = max(self._last_hit_index, row.index)
            for r in rows:
                if lo <= r.index <= hi:
                    r.set_selected(self._drag_state)
            self._last_hit_index = row.index
            self.on_change()

    def _on_release(self, event):
        self._dragging = False
        self._last_hit_index = -1


# ============================================================================
# MAIN APPLICATION
# ============================================================================

class PlaylistCreatorApp(ctk.CTk):
    GRID_MIN_TILE = 120   # pixels — tile won't shrink below this
    GRID_MAX_TILE = 200   # pixels — tile won't grow beyond this

    def __init__(self):
        super().__init__()

        self.title("RedlineAuto & RoyalRenderings PlaylistCreator")
        self.geometry("1200x800")
        self.minsize(1000, 650)
        self.resizable(True, True)

        try:
            # Try RR icon first, then RLA
            for ico in ("RR-Icon.ico", "RLA-PlaylistCreator.ico"):
                ico_path = resource_path(ico)
                if os.path.isfile(ico_path):
                    self.iconbitmap(ico_path)
                    break
        except Exception:
            pass

        self.settings = load_settings()

        ctk.set_appearance_mode(self.settings["appearance"])
        ctk.set_default_color_theme("blue")
        self.configure(fg_color=("#f0f0f0", "#0a0a0a"))

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.songs = []
        self.song_rows = []
        self.grid_tiles = []
        self.logo_image = None
        self.playlist_title = ""
        self._stop_loading = False
        self._stop_requested = False
        self._loading = False
        self._known_ids = set()  # video_id dedup
        self._batch_idx_to_tile = {}  # Maps fetch song_idx → grid_tiles index
        self._grid_cols = 4          # Current column count (recalculated on resize)
        self._grid_resize_after = None  # Debounce handle
        self._log_buffer = ""          # Persistent log text
        self._log_window = None        # LogWindow instance

        self._build_header()
        self._build_main_content()
        self._build_footer()
        self._apply_appearance(self.settings["appearance"])

        # Install deps + ffmpeg in background on launch
        def _startup_deps():
            # ── Auto-update check (frozen exe only) ─────────────
            if getattr(sys, 'frozen', False) and _should_check_update():
                self._log("Checking for updates...")
                _mark_update_checked()
                update_info = check_for_update(log_callback=self._log)

                if update_info:
                    user_response = [None]
                    event = threading.Event()

                    def _ask():
                        result = messagebox.askyesno(
                            "Update Available",
                            f"A new version ({update_info['version']}) is available.\n\n"
                            f"You are running v{__version__}.\n\n"
                            f"Would you like to update now?",
                            parent=self
                        )
                        user_response[0] = result
                        event.set()

                    self.after(0, _ask)
                    event.wait(timeout=120)

                    if user_response[0]:
                        self.after(0, lambda: (
                            self._setup_label.configure(text="Downloading update..."),
                            self._setup_label.pack(pady=(0, 4))
                        ))

                        def _dl_log(msg):
                            self._log(msg)
                            self.after(0, lambda m=msg: self._setup_label.configure(text=m))

                        installer_path = download_update(
                            update_info["download_url"],
                            update_info["asset_name"],
                            log_callback=_dl_log
                        )

                        if installer_path:
                            self._log(f"Launching installer: {installer_path}")
                            subprocess.Popen(
                                [str(installer_path), '/SILENT', '/CLOSEAPPLICATIONS'],
                                startupinfo=_hidden_startupinfo()
                            )
                            self.after(500, lambda: self.destroy())
                            return
                        else:
                            self.after(0, lambda: self._setup_label.pack_forget())
                            self._log("Update download failed — continuing normally")

            # ── Dependency check ────────────────────────────────
            needs_setup = False

            # Check if anything needs installing
            modules = ["requests", "yt_dlp", "PIL", "mutagen", "numpy"]
            missing = [m for m in modules if importlib.util.find_spec(m) is None]
            needs_ffmpeg = get_ffmpeg_path() is None

            if missing or needs_ffmpeg:
                needs_setup = True
                self.after(0, lambda: (
                    self._setup_label.configure(text="Setting up for first use..."),
                    self._setup_label.pack(pady=(0, 4))
                ))

            def _setup_log(msg):
                self._log(msg)
                if needs_setup:
                    self.after(0, lambda m=msg: self._setup_label.configure(text=m))

            # Install missing packages
            check_and_install_dependencies(log_callback=_setup_log)

            # Auto-install ffmpeg if not found
            if needs_ffmpeg:
                ensure_ffmpeg(log_callback=_setup_log)

            # Hide setup label
            if needs_setup:
                self.after(0, lambda: self._setup_label.pack_forget())

        threading.Thread(target=_startup_deps, daemon=True).start()

    # ── Header ──────────────────────────────────────────────

    def _build_header(self):
        self._header = ctk.CTkFrame(self, fg_color=("#f0f0f0", "#0a0a0a"), corner_radius=0, height=110)
        self._header.grid(row=0, column=0, sticky="ew")
        self._header.grid_propagate(False)
        header = self._header

        top_bar = ctk.CTkFrame(header, fg_color="transparent")
        top_bar.pack(fill="x", padx=15, pady=(8, 0))
        ctk.CTkButton(
            top_bar, text="Settings", command=self._open_settings,
            width=100, height=28, font=ctk.CTkFont(size=11),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color=("#d0d0d0", "#2e2e2e"), corner_radius=6,
            text_color=("#333333", "#d0d0d0")
        ).pack(side="right")

        ctk.CTkButton(
            top_bar, text="Log", command=self._open_log,
            width=60, height=28, font=ctk.CTkFont(size=11),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color=("#d0d0d0", "#2e2e2e"), corner_radius=6,
            text_color=("#333333", "#d0d0d0")
        ).pack(side="right", padx=(0, 6))

        self.logo_label = ctk.CTkLabel(
            header, text="RA x RR",
            font=ctk.CTkFont(size=24, weight="bold"), text_color=("#111111", "#ffffff")
        )
        self.logo_label.pack(pady=(0, 2))

        def _load_logo():
            try:
                img = Image.open(resource_path("RA-x-RR.png"))
                h = 55
                w = int(h * (img.width / img.height))
                img = img.resize((w, h), Image.Resampling.LANCZOS)
                self.logo_image = ctk.CTkImage(light_image=img, dark_image=img, size=(w, h))
                self.logo_label.configure(image=self.logo_image, text="")
            except Exception:
                pass
        threading.Thread(target=_load_logo, daemon=True).start()

        ctk.CTkLabel(
            header, text="BeamNG Playlist Creator",
            font=ctk.CTkFont(size=16, weight="bold"), text_color=("#111111", "#ffffff")
        ).pack(pady=(0, 2))

        ctk.CTkLabel(
            header, text="-For RedlineAuto and RoyalRenderings mods only-",
            font=ctk.CTkFont(size=10), text_color=("#808080", "#606060")
        ).pack(pady=(0, 8))

        # Setup status label — bold, visible only during first-time installs
        self._setup_label = ctk.CTkLabel(
            header, text="",
            font=ctk.CTkFont(size=12, weight="bold"), text_color="#ff5959"
        )
        # Not packed yet — shown by _startup_deps if needed

    # ── Main Content (Grid + List side by side) ─────────────

    def _build_main_content(self):
        self.content = ctk.CTkFrame(self, fg_color="transparent")
        self.content.grid(row=1, column=0, sticky="nsew", padx=20, pady=15)
        self.content.grid_columnconfigure(0, weight=5)   # Grid panel - much wider
        self.content.grid_columnconfigure(1, weight=2)   # List panel - compact
        self.content.grid_rowconfigure(0, weight=1)

        # ── Left: Song Grid ──
        self._grid_panel = ctk.CTkFrame(self.content, fg_color=("#ffffff", "#111111"), corner_radius=10)
        self._grid_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        grid_panel = self._grid_panel
        grid_panel.grid_rowconfigure(1, weight=1)
        grid_panel.grid_columnconfigure(0, weight=1)

        grid_header = ctk.CTkFrame(grid_panel, fg_color="transparent", height=40)
        grid_header.grid(row=0, column=0, sticky="ew", padx=15, pady=(12, 0))
        grid_header.grid_propagate(False)

        ctk.CTkLabel(
            grid_header, text="Song Grid",
            font=ctk.CTkFont(size=14, weight="bold"), text_color=("#1a1a1a", "#e0e0e0")
        ).pack(side="left")

        self.grid_count_label = ctk.CTkLabel(
            grid_header, text="No songs loaded",
            font=ctk.CTkFont(size=11), text_color=("#888888", "#505050")
        )
        self.grid_count_label.pack(side="right")

        self.grid_scroll = ctk.CTkScrollableFrame(
            grid_panel, fg_color="transparent",
            scrollbar_button_color=("#d0d0d0", "#1e1e1e"),
            scrollbar_button_hover_color=("#d0d0d0", "#2e2e2e")
        )
        self.grid_scroll.grid(row=1, column=0, sticky="nsew", padx=8, pady=(8, 12))

        # Add Song tile (always visible as first tile)
        self.add_song_tile = AddSongTile(
            self.grid_scroll, on_click=self._show_add_song_popup, tile_size=128
        )
        self.add_song_tile.grid(row=0, column=0, padx=4, pady=4, sticky="nsew")

        # Configure initial columns (recalculated dynamically on resize)
        for c in range(self._grid_cols):
            self.grid_scroll.grid_columnconfigure(c, weight=1, uniform="tile")

        # Recalculate columns whenever the grid panel resizes
        self.grid_scroll.bind("<Configure>", self._schedule_grid_resize, add="+")

        # ── Right: Song List ──
        self._list_panel = ctk.CTkFrame(self.content, fg_color=("#ffffff", "#111111"), corner_radius=10)
        self._list_panel.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        list_panel = self._list_panel
        list_panel.grid_rowconfigure(1, weight=1)
        list_panel.grid_columnconfigure(0, weight=1)

        list_header = ctk.CTkFrame(list_panel, fg_color="transparent", height=40)
        list_header.grid(row=0, column=0, sticky="ew", padx=15, pady=(12, 0))
        list_header.grid_propagate(False)

        ctk.CTkLabel(
            list_header, text="Song List",
            font=ctk.CTkFont(size=14, weight="bold"), text_color=("#1a1a1a", "#e0e0e0")
        ).pack(side="left")

        btn_row = ctk.CTkFrame(list_header, fg_color="transparent")
        btn_row.pack(side="right")

        ctk.CTkButton(
            btn_row, text="All", command=self._select_all,
            width=38, height=26, font=ctk.CTkFont(size=10),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color=("#d0d0d0", "#2e2e2e"), corner_radius=4,
            text_color=("#333333", "#d0d0d0")
        ).pack(side="left", padx=(0, 3))

        ctk.CTkButton(
            btn_row, text="None", command=self._deselect_all,
            width=42, height=26, font=ctk.CTkFont(size=10),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color=("#d0d0d0", "#2e2e2e"), corner_radius=4,
            text_color=("#333333", "#d0d0d0")
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_row, text="Remove Unselected", command=self._remove_unselected,
            width=120, height=26, font=ctk.CTkFont(size=10),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color="#7f1d1d", corner_radius=4,
            text_color=("#555555", "#808080")
        ).pack(side="left", padx=(0, 3))

        ctk.CTkButton(
            btn_row, text="Remove Selected", command=self._remove_selected,
            width=110, height=26, font=ctk.CTkFont(size=10),
            fg_color=("#e0e0e0", "#1e1e1e"), hover_color="#7f1d1d", corner_radius=4,
            text_color=("#555555", "#808080")
        ).pack(side="left")

        self.list_scroll = ctk.CTkScrollableFrame(
            list_panel, fg_color="transparent",
            scrollbar_button_color=("#d0d0d0", "#1e1e1e"),
            scrollbar_button_hover_color=("#d0d0d0", "#2e2e2e")
        )
        self.list_scroll.grid(row=1, column=0, sticky="nsew", padx=4, pady=(8, 4))
        self.list_scroll.grid_columnconfigure(0, weight=1)

        # Add Song row (always visible as first row)
        self.add_song_row = AddSongRow(
            self.list_scroll, on_click=self._show_add_song_popup
        )
        self.add_song_row.pack(fill="x", pady=1)

        # Status / progress
        self.status_label = ctk.CTkLabel(
            list_panel, text="",
            font=ctk.CTkFont(size=11), text_color=("#888888", "#505050"),
            anchor="w"
        )
        self.status_label.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 2))

        # Bottom bar
        list_bottom = ctk.CTkFrame(list_panel, fg_color="transparent", height=35)
        list_bottom.grid(row=3, column=0, sticky="ew", padx=15, pady=(0, 10))
        list_bottom.grid_propagate(False)

        self.selection_label = ctk.CTkLabel(
            list_bottom, text="0/0 selected",
            font=ctk.CTkFont(size=11), text_color=("#888888", "#505050")
        )
        self.selection_label.pack(side="left", pady=4)

        # Drag-select manager
        self.drag_manager = DragSelectManager(
            self.list_scroll,
            get_rows_fn=lambda: self.song_rows,
            on_change_fn=self._update_selection_count
        )

    # ── Grid resize ─────────────────────────────────────────

    def _schedule_grid_resize(self, _event=None):
        """Debounce <Configure> events so we only reflow after resizing stops"""
        if self._grid_resize_after:
            self.after_cancel(self._grid_resize_after)
        self._grid_resize_after = self.after(120, self._on_grid_resize)

    def _on_grid_resize(self):
        self._grid_resize_after = None
        try:
            width = self.grid_scroll.winfo_width() - 16  # subtract ~scrollbar width
        except Exception:
            return
        if width < 50:
            return

        # Ceiling division: minimum columns so no tile exceeds GRID_MAX_TILE
        min_cols = max(1, (width + self.GRID_MAX_TILE - 1) // self.GRID_MAX_TILE)
        # Floor division: maximum columns so no tile is narrower than GRID_MIN_TILE
        max_cols = max(1, width // self.GRID_MIN_TILE)
        # Keep as close to 4 columns as the bounds allow
        new_cols = max(min_cols, min(max_cols, 4))

        if new_cols != self._grid_cols:
            self._grid_cols = new_cols
            self._reflow_grid()

    def _reflow_grid(self):
        """Re-configure grid columns and re-place all tiles at the new column count"""
        cols = self._grid_cols

        # Clear old column configs (up to 20 previous columns)
        for c in range(20):
            try:
                self.grid_scroll.grid_columnconfigure(c, weight=0, minsize=0, uniform="")
            except Exception:
                break

        # Apply new column configs
        for c in range(cols):
            self.grid_scroll.grid_columnconfigure(c, weight=1, uniform="tile")

        # Re-place the Add Song tile at col 0
        self.add_song_tile.grid(row=0, column=0, padx=4, pady=4, sticky="nsew")

        # Re-place all song tiles
        for i, tile in enumerate(self.grid_tiles):
            grid_pos = i + 1  # offset by 1 for the AddSongTile
            tile.grid(row=grid_pos // cols, column=grid_pos % cols,
                      padx=4, pady=4, sticky="nsew")

    # ── Footer ──────────────────────────────────────────────

    def _build_footer(self):
        footer = ctk.CTkFrame(self, fg_color="transparent", height=70)
        footer.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 15))
        footer.grid_columnconfigure(0, weight=1)

        self._action_bar = ctk.CTkFrame(footer, fg_color=("#ffffff", "#111111"), corner_radius=10, height=55)
        self._action_bar.pack(fill="x")
        self._action_bar.pack_propagate(False)
        action_bar = self._action_bar

        inner = ctk.CTkFrame(action_bar, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=12, pady=8)
        inner.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            inner, text="Pack Name",
            font=ctk.CTkFont(size=12, weight="bold"), text_color=("#606060", "#909090")
        ).grid(row=0, column=0, padx=(4, 10))

        self.name_entry = ctk.CTkEntry(
            inner, placeholder_text="Enter radiopack name...",
            height=36, font=ctk.CTkFont(size=12),
            border_width=0, corner_radius=6, fg_color=("#e8e8e8", "#1a1a1a")
        )
        self.name_entry.grid(row=0, column=1, sticky="ew")

        self.create_btn = ctk.CTkButton(
            inner, text="Create Radiopack", command=self._create_radiopack,
            width=160, height=36, font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#16a34a", hover_color="#15803d", corner_radius=8
        )
        self.create_btn.grid(row=0, column=2, padx=(12, 4))

        self.stop_btn = ctk.CTkButton(
            inner, text="Stop", command=self._stop_creation,
            width=80, height=36, font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#dc2626", hover_color="#b91c1c", corner_radius=8
        )
        # Stop button starts hidden — shown by _lock_ui()

        # Progress bar and label (below action bar, initially hidden)
        self._progress_frame = ctk.CTkFrame(footer, fg_color="transparent")
        self.progress_bar = ctk.CTkProgressBar(
            self._progress_frame, height=8, corner_radius=4,
            fg_color=("#e0e0e0", "#1a1a1a"), progress_color="#16a34a"
        )
        self.progress_bar.pack(fill="x", padx=4, pady=(4, 2))
        self.progress_bar.set(0)
        self.progress_label = ctk.CTkLabel(
            self._progress_frame, text="",
            font=ctk.CTkFont(size=11), text_color=("#606060", "#909090")
        )
        self.progress_label.pack(anchor="w", padx=4)
        # Progress frame starts hidden

        ctk.CTkLabel(
            footer, text="Made by RedlineAuto & RoyalRenderings 2026",
            font=ctk.CTkFont(size=10), text_color=("#bbbbbb", "#383838")
        ).pack(pady=(6, 0))

    # ── Actions ─────────────────────────────────────────────

    def _open_settings(self):
        if hasattr(self, "_settings_dialog") and self._settings_dialog is not None:
            try:
                if self._settings_dialog.winfo_exists():
                    # Force window visible and to front after theme switches
                    self._settings_dialog.deiconify()
                    self._settings_dialog.attributes("-topmost", True)
                    self._settings_dialog.after(100, lambda: self._settings_dialog.attributes("-topmost", False))
                    self._settings_dialog.focus_force()
                    return
            except Exception:
                pass
            self._settings_dialog = None
        self._settings_dialog = SettingsDialog(self, self.settings, on_save=self._on_settings_changed)

    def _open_log(self):
        if self._log_window is not None:
            try:
                if self._log_window.winfo_exists():
                    self._log_window.deiconify()
                    self._log_window.attributes("-topmost", True)
                    self._log_window.after(100, lambda: self._log_window.attributes("-topmost", False))
                    self._log_window.focus_force()
                    return
            except Exception:
                pass
            self._log_window = None
        self._log_window = LogWindow(self)

    def _log(self, text):
        """Central logging — appends to buffer and log window. Thread-safe."""
        def _do():
            self._log_buffer += text + "\n"
            if self._log_window is not None:
                try:
                    if self._log_window.winfo_exists():
                        self._log_window.append(text)
                except Exception:
                    pass
        self.after(0, _do)

    def _on_settings_changed(self, settings):
        self.settings = settings
        # Defer appearance change to avoid conflicts during CTk's internal update cycle
        self.after(50, lambda: self._apply_appearance(settings["appearance"]))

    def _apply_appearance(self, mode):
        ctk.set_appearance_mode(mode)
        # Force-refresh root and major containers (CTk root doesn't auto-update tuples)
        root_bg = ("#f0f0f0", "#0a0a0a")
        panel_bg = ("#ffffff", "#111111")
        self.configure(fg_color=root_bg)
        self._header.configure(fg_color=root_bg)
        for panel in (self._grid_panel, self._list_panel, self._action_bar):
            panel.configure(fg_color=panel_bg)

    def _is_playlist_url(self, url):
        """Detect if URL is a playlist vs single song"""
        return 'list=' in url and '/playlist' in url

    def _show_add_song_popup(self):
        """Open the Add Song / Playlist dialog"""
        if self._loading:
            messagebox.showinfo("Busy", "A playlist is currently loading. Please wait or cancel first.")
            return
        AddSongDialog(self, on_submit=self._handle_add_url)

    def _handle_add_url(self, url):
        """Route to playlist or single song handler"""
        if self._is_playlist_url(url):
            self._add_playlist(url)
        else:
            self._add_single_song(url)

    # ── Add Playlist (additive — stacks with existing songs) ──

    def _add_playlist(self, url):
        """Load a playlist and APPEND songs to the existing list"""
        self._stop_loading = False
        self._loading = True
        self._cover_batch_offset = len(self.grid_tiles)  # Track where this batch starts
        self._batch_idx_to_tile = {}  # Maps fetch song_idx → grid_tiles index
        self.status_label.configure(text="Loading playlist...")

        def _log_and_status(msg):
            self.after(0, lambda m=msg: self.status_label.configure(text=m))
            self._log(msg)

        def _do_load():
            try:
                check_and_install_dependencies(log_callback=_log_and_status)

                playlist_title, songs = fetch_playlist_info(
                    url,
                    progress_callback=lambda kind, val: self.after(0, lambda k=kind, v=val: self._on_progress(k, v)),
                    stop_check=lambda: self._stop_loading
                )
                self.playlist_title = playlist_title
                self._log(f"Playlist loaded: {playlist_title} ({len(songs)} songs)")
                self.after(0, self._on_load_complete)

            except Exception as e:
                self._log(f"ERROR loading playlist: {str(e)[:200]}")
                self.after(0, lambda: self._on_load_error(str(e)))

        threading.Thread(target=_do_load, daemon=True).start()

    def _on_progress(self, kind, value):
        """Handle progressive loading from the fetch thread"""
        if kind == "status":
            self.status_label.configure(text=value)

        elif kind == "total":
            pass  # We don't show a total since songs stack

        elif kind == "song":
            idx, song = value

            # ── Dedup by video_id ──
            vid = song.get('video_id', '')
            if vid and vid in self._known_ids:
                return  # Skip duplicate
            if vid:
                self._known_ids.add(vid)

            self.songs.append(song)
            count = len(self.songs)

            # Add tile to grid (offset by 1 for the AddSongTile at position 0)
            cols = self._grid_cols
            grid_pos = count
            tile = SongTile(self.grid_scroll, song, tile_size=128)
            tile.grid(row=grid_pos // cols, column=grid_pos % cols, padx=4, pady=4, sticky="nsew")
            self.grid_tiles.append(tile)

            # Track batch index → tile index for cover updates
            self._batch_idx_to_tile[idx] = len(self.grid_tiles) - 1

            # Add row to list
            row = SongRow(
                self.list_scroll, index=count, song_data=song,
                on_toggle=self._update_selection_count,
                on_remove=self._remove_song_by_index
            )
            row.pack(fill="x", pady=1)
            self.song_rows.append(row)

            self.grid_count_label.configure(text=f"{count} songs")
            self.status_label.configure(text=f"Loaded {count} songs")
            self._update_selection_count()

        elif kind == "cover":
            song_idx, img = value
            tile_idx = self._batch_idx_to_tile.get(song_idx)
            if tile_idx is not None and 0 <= tile_idx < len(self.grid_tiles):
                tile = self.grid_tiles[tile_idx]
                art_size = 128 - 16
                resized = img.resize((art_size, art_size), Image.Resampling.LANCZOS)
                tile._ctk_image = ctk.CTkImage(light_image=resized, dark_image=resized, size=(art_size, art_size))
                for child in tile.winfo_children():
                    if isinstance(child, ctk.CTkLabel) and child.cget("text") == "":
                        child.configure(image=tile._ctk_image)
                        break

    def _on_load_complete(self):
        self._loading = False
        count = len(self.songs)
        self.status_label.configure(text=f"Loaded {count} songs total  (latest: {self.playlist_title})")
        self.grid_count_label.configure(text=f"{count} songs")

    def _on_load_error(self, error_msg):
        self._loading = False
        self.status_label.configure(text=f"Error: {error_msg[:120]}")
        messagebox.showerror("Load Error", f"Failed to load playlist:\n\n{error_msg[:300]}")

    def _clear_panels(self):
        for tile in self.grid_tiles:
            tile.destroy()
        self.grid_tiles.clear()

        for row in self.song_rows:
            row.destroy()
        self.song_rows.clear()

        self.songs.clear()
        self._known_ids.clear()
        self._batch_idx_to_tile.clear()

    def _update_selection_count(self):
        total = len(self.song_rows)
        selected = sum(1 for r in self.song_rows if r.selected.get())
        self.selection_label.configure(text=f"{selected}/{total} selected")

    def _select_all(self):
        for row in self.song_rows:
            row.set_selected(True)
        self._update_selection_count()

    def _deselect_all(self):
        for row in self.song_rows:
            row.set_selected(False)
        self._update_selection_count()

    # ── Remove Songs ────────────────────────────────────────

    def _remove_song_by_index(self, row_index):
        """Remove a single song by its row index (1-based)"""
        idx = row_index - 1  # Convert to 0-based
        if 0 <= idx < len(self.songs):
            # Remove from dedup set
            vid = self.songs[idx].get('video_id', '')
            self._known_ids.discard(vid)

            self.songs.pop(idx)
            self._rebuild_panels()
            self.status_label.configure(text=f"Removed song #{row_index}")

    def _remove_selected(self):
        """Remove all currently selected songs"""
        if not self.song_rows:
            return
        selected_count = sum(1 for r in self.song_rows if r.selected.get())
        if selected_count == 0:
            return

        keep = []
        for song, row in zip(self.songs, self.song_rows):
            if row.selected.get():
                self._known_ids.discard(song.get('video_id', ''))
            else:
                keep.append(song)

        self.songs = keep
        self._rebuild_panels()
        self.status_label.configure(text=f"Removed {selected_count} selected songs")

    def _remove_unselected(self):
        """Remove all currently unselected songs"""
        if not self.song_rows:
            return
        unselected_count = sum(1 for r in self.song_rows if not r.selected.get())
        if unselected_count == 0:
            return

        keep = []
        for song, row in zip(self.songs, self.song_rows):
            if not row.selected.get():
                self._known_ids.discard(song.get('video_id', ''))
            else:
                keep.append(song)

        self.songs = keep
        self._rebuild_panels()
        self.status_label.configure(text=f"Removed {unselected_count} unselected songs")

    def _rebuild_panels(self):
        """Destroy all tiles/rows and recreate from self.songs"""
        # Destroy existing
        for tile in self.grid_tiles:
            tile.destroy()
        self.grid_tiles.clear()

        for row in self.song_rows:
            row.destroy()
        self.song_rows.clear()

        self._batch_idx_to_tile.clear()

        # Recreate
        cols = self._grid_cols
        for i, song in enumerate(self.songs):
            grid_pos = i + 1  # offset for add tile
            tile = SongTile(self.grid_scroll, song, tile_size=128)
            tile.grid(row=grid_pos // cols, column=grid_pos % cols, padx=4, pady=4, sticky="nsew")
            self.grid_tiles.append(tile)

            row = SongRow(
                self.list_scroll, index=i + 1, song_data=song,
                on_toggle=self._update_selection_count,
                on_remove=self._remove_song_by_index
            )
            row.pack(fill="x", pady=1)
            self.song_rows.append(row)

        count = len(self.songs)
        self.grid_count_label.configure(text=f"{count} songs" if count else "No songs loaded")
        self._update_selection_count()

        # If the list is now empty, scroll back to top so the Add Song row is visible
        if not self.songs:
            try:
                self.list_scroll._parent_canvas.yview_moveto(0)
            except Exception:
                pass

    def _create_radiopack(self):
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showerror("Missing Name", "Please enter a radiopack name!")
            return

        selected = [s for s, r in zip(self.songs, self.song_rows) if r.selected.get()]
        if not selected:
            messagebox.showerror("No Songs", "Please select at least one song!")
            return

        self._stop_requested = False
        self._lock_ui()

        def _worker():
            try:
                self._run_radiopack_creation(name, selected)
            except Exception as e:
                self._update_progress(f"ERROR: {str(e)[:120]}")
            finally:
                self.after(0, self._unlock_ui)

        threading.Thread(target=_worker, daemon=True).start()

    def _run_radiopack_creation(self, playlist_name, selected_songs):
        """Main backend orchestrator — runs on background thread"""
        total_songs = len(selected_songs)
        audio_format = self.settings.get("audio_format", "mp3")
        audio_quality = self.settings.get("audio_quality", "192")
        output_folder = self.settings.get("output_folder", _default_output_folder())

        log = self._update_progress

        # ── 1. Ensure dependencies ──
        log("Checking dependencies...")
        self._set_progress(0)
        check_and_install_dependencies(log_callback=log)

        if self._stop_requested:
            log("Cancelled.")
            return

        if not ensure_ffmpeg(log_callback=log, stop_check=lambda: self._stop_requested):
            log("ERROR: ffmpeg is required but could not be installed")
            return

        if self._stop_requested:
            log("Cancelled.")
            return

        # ── 2. Create directory structure ──
        clean_name = clean_filename(playlist_name)
        output_dir = Path(output_folder)

        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            log(f"ERROR: Cannot write to output folder: {str(e)[:100]}")
            return

        final_name, base_path = find_available_name(clean_name, output_dir, log_callback=log)

        music_dir = base_path / "art" / "sound" / "music"
        songs_dir = base_path / "vehicles" / "common" / "lua" / "songs"
        covers_dir = base_path / "vehicles" / "common" / "album_covers"

        for d in [music_dir, songs_dir, covers_dir]:
            d.mkdir(parents=True, exist_ok=True)

        log(f"Created radiopack structure: {final_name}")

        # ── 3. Save album covers from in-memory PIL images ──
        log("Saving album covers...")
        cover_images = []
        for song in selected_songs:
            clean_title = clean_filename(song['title'])
            cover = song.get('cover')
            if cover and isinstance(cover, Image.Image):
                cover_images.append(cover)
                try:
                    # Convert to RGB if needed (some PNGs are RGBA)
                    save_img = cover.convert('RGB') if cover.mode != 'RGB' else cover
                    save_img.save(covers_dir / f"{clean_title}_art.png")
                except Exception:
                    pass

        # Create playlist cover (2x2 grid)
        lua_var = playlist_name.lower().replace(' ', '_').replace('-', '_')
        cover_out = covers_dir / f"{lua_var}_playlist.png"
        create_playlist_cover_from_pil(cover_images, cover_out)

        if self._stop_requested:
            shutil.rmtree(base_path, ignore_errors=True)
            log("Cancelled. Cleaned up files.")
            return

        # ── 4. Download & process audio per song ──
        import yt_dlp

        format_chain = [
            f'bestaudio[ext=m4a]/bestaudio[ext=mp3]/bestaudio/best',
            'bestaudio/best',
            'best',
        ]

        skipped = 0
        successful = 0

        for idx, song in enumerate(selected_songs, 1):
            if self._stop_requested:
                log("Cancelled by user.")
                break

            title = song['title']
            video_id = song.get('video_id', '')
            clean_title = clean_filename(title)

            if not video_id:
                skipped += 1
                log(f"[{idx}/{total_songs}] Skipped: {title[:40]} -- no video ID")
                continue

            video_url = f"https://www.youtube.com/watch?v={video_id}"
            log(f"[{idx}/{total_songs}] Downloading: {title[:50]}...")
            self._set_progress(idx / (total_songs + 2))  # +2 for Lua + ZIP steps

            # Snapshot files before download
            before_files = set(f.name for f in music_dir.glob(f"*.{audio_format}")
                               if not f.stem.endswith('_bass'))

            # Try download with format fallback + retries
            download_success = False
            last_error = None

            for fmt in format_chain:
                if download_success:
                    break
                dl_opts = {
                    **_ydl_opts(),
                    'format': fmt,
                    'postprocessors': [{
                        'key': 'FFmpegExtractAudio',
                        'preferredcodec': audio_format,
                        'preferredquality': audio_quality,
                    }],
                    'outtmpl': str(music_dir / '%(title)s.%(ext)s'),
                    'prefer_ffmpeg': True,
                    'keepvideo': False,
                    'ignoreerrors': False,
                    'extract_flat': False,
                    'noplaylist': True,
                }

                for attempt in range(3):
                    if self._stop_requested:
                        break
                    try:
                        with yt_dlp.YoutubeDL(dl_opts) as dl:
                            dl.download([video_url])
                        download_success = True
                        break
                    except Exception as e:
                        last_error = str(e)
                        if _is_permanent_error(last_error):
                            break
                        if attempt < 2:
                            time.sleep(1.5 * (attempt + 1))

            if self._stop_requested:
                break

            if not download_success:
                skipped += 1
                reason = _strip_ansi(last_error or "Unknown").strip()
                log(f"[{idx}/{total_songs}] Failed: {title[:40]} -- {reason}")
                continue

            # Find the new file via directory diff
            time.sleep(0.3)
            after_files = set(f.name for f in music_dir.glob(f"*.{audio_format}")
                              if not f.stem.endswith('_bass'))
            new_files = after_files - before_files

            if not new_files:
                time.sleep(0.5)
                after_files = set(f.name for f in music_dir.glob(f"*.{audio_format}")
                                  if not f.stem.endswith('_bass'))
                new_files = after_files - before_files

            main_path = None
            if new_files:
                main_path = music_dir / list(new_files)[0]
                # Rename to clean filename
                new_path = music_dir / f"{clean_title}.{audio_format}"
                try:
                    if main_path != new_path:
                        if new_path.exists():
                            new_path = music_dir / f"{clean_title}_{idx}.{audio_format}"
                        main_path.rename(new_path)
                        main_path = new_path
                except Exception:
                    pass

            if not main_path or not main_path.exists():
                skipped += 1
                log(f"[{idx}/{total_songs}] File not found after download: {title[:40]}")
                continue

            if main_path.stat().st_size < 1024:
                skipped += 1
                main_path.unlink(missing_ok=True)
                log(f"[{idx}/{total_songs}] Downloaded file is empty: {title[:40]}")
                continue

            successful += 1
            clean_stem = main_path.stem

            # Create bass track
            if self._stop_requested:
                break
            log(f"[{idx}/{total_songs}] Creating bass track...")
            bass_ext = audio_format  # bass track in same format
            create_bass_track(main_path, music_dir / f"{clean_stem}_bass.{bass_ext}", log)

            # Generate amplitude data
            if self._stop_requested:
                break
            log(f"[{idx}/{total_songs}] Generating amplitude data...")
            generate_amplitude_data(main_path, music_dir / f"{clean_stem}_data.json", log)

        # Clean up temp files (non-audio, non-json leftovers)
        for leftover in music_dir.iterdir():
            if leftover.suffix.lower() not in (f'.{audio_format}', '.json'):
                leftover.unlink(missing_ok=True)

        if self._stop_requested:
            shutil.rmtree(base_path, ignore_errors=True)
            log("Cancelled. Cleaned up files.")
            return

        if successful == 0:
            shutil.rmtree(base_path, ignore_errors=True)
            log(f"ERROR: No songs downloaded successfully ({skipped} skipped)")
            return

        log(f"Downloaded {successful}/{total_songs} songs")

        # ── 5. Generate Lua playlist ──
        log("Generating Lua playlist...")
        self._set_progress((total_songs + 1) / (total_songs + 2))
        generate_lua_playlist(music_dir, songs_dir, covers_dir, final_name,
                              selected_songs, audio_format, log)

        # ── 6. Create ZIP archive ──
        self._set_progress((total_songs + 1.5) / (total_songs + 2))
        zip_path = create_zip_archive(base_path, final_name, log)

        # ── 7. Clean up temp directory (keep ZIP) ──
        if zip_path:
            try:
                shutil.rmtree(base_path)
                log("Cleaned up temp files")
            except Exception:
                pass

        # ── Done ──
        self._set_progress(1.0)
        if skipped > 0:
            log(f"Done! {successful}/{total_songs} songs ({skipped} skipped)")
        else:
            log(f"Done! Radiopack created with {successful} songs")

        result_path = zip_path or base_path
        self.after(0, lambda: messagebox.showinfo(
            "Radiopack Created",
            f"'{final_name}' created successfully!\n\n"
            f"{successful} songs processed\n"
            f"Saved to:\n{result_path}"
        ))

    def _stop_creation(self):
        """Signal the background thread to stop"""
        self._stop_requested = True
        self._update_progress("Stopping...")

    def _lock_ui(self):
        """Disable UI during radiopack creation"""
        self.create_btn.configure(state="disabled", text="Creating...")
        self.name_entry.configure(state="disabled")
        self.stop_btn.grid(row=0, column=3, padx=(8, 4))
        self._progress_frame.pack(fill="x", pady=(4, 0))
        self.progress_bar.set(0)
        self.progress_label.configure(text="Starting...")

    def _unlock_ui(self):
        """Re-enable UI after radiopack creation"""
        self.create_btn.configure(state="normal", text="Create Radiopack")
        self.name_entry.configure(state="normal")
        self.stop_btn.grid_remove()
        # Keep progress visible briefly so user can see final status

    def _update_progress(self, text):
        """Thread-safe progress text update — also routes to log window"""
        self.after(0, lambda: self.progress_label.configure(text=text))
        self._log(text)

    def _set_progress(self, fraction):
        """Thread-safe progress bar update (0.0 to 1.0)"""
        self.after(0, lambda: self.progress_bar.set(min(1.0, max(0.0, fraction))))

    # ── Add Single Song ─────────────────────────────────────

    def _add_single_song(self, url):
        """Fetch a single song and append it to the grid + list"""
        self.status_label.configure(text="Adding song...")

        def _do_fetch():
            try:
                check_and_install_dependencies()
                song = fetch_single_song(url)
                self.after(0, lambda: self._append_song(song))
            except Exception as e:
                self.after(0, lambda: self._on_add_song_error(str(e)))

        threading.Thread(target=_do_fetch, daemon=True).start()

    def _append_song(self, song):
        """Append a single song to both panels (with dedup)"""
        vid = song.get('video_id', '')
        if vid and vid in self._known_ids:
            self.status_label.configure(text=f"Skipped duplicate: {song['title']}")
            self._log(f"Skipped duplicate: {song['title']}")
            return
        if vid:
            self._known_ids.add(vid)

        self.songs.append(song)
        count = len(self.songs)

        # Add tile (offset by 1 for add tile)
        cols = self._grid_cols
        grid_pos = count
        tile = SongTile(self.grid_scroll, song, tile_size=128)
        tile.grid(row=grid_pos // cols, column=grid_pos % cols, padx=4, pady=4, sticky="nsew")
        self.grid_tiles.append(tile)

        # Add row
        row = SongRow(
            self.list_scroll, index=count, song_data=song,
            on_toggle=self._update_selection_count,
            on_remove=self._remove_song_by_index
        )
        row.pack(fill="x", pady=1)
        self.song_rows.append(row)

        self.grid_count_label.configure(text=f"{count} songs")
        self.status_label.configure(text=f"Added: {song['title']} -- {song['artist']}")
        self._log(f"Added: {song['title']} -- {song['artist']}")
        self._update_selection_count()

    def _on_add_song_error(self, error_msg):
        self.status_label.configure(text=f"Error adding song: {error_msg[:100]}")
        self._log(f"ERROR adding song: {error_msg[:200]}")
        messagebox.showerror("Add Song Error", f"Failed to add song:\n\n{error_msg[:300]}")


# ============================================================================
# ENTRY POINT
# ============================================================================

def main():
    app = PlaylistCreatorApp()
    app.mainloop()

if __name__ == "__main__":
    main()