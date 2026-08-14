"""SongWeaver - a 20-pad local soundboard for Windows.

Assign audio files (mp3, wav, wma, m4a) from your library to pads and
trigger them by click or keyboard. Playback goes through the Windows Media
Player COM control (pywin32), which has far broader codec support (notably
M4A/AAC) than the legacy MCI API this app used to rely on.
"""

import json
import os
import sys
import time
import uuid
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

import pythoncom
import win32com.client

BASE_PAD_COUNT = 20
MAX_PAD_COUNT = 100
PAD_INCREMENT = 5  # one row
COLS = 5
KEYS = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0',
        'Q', 'W', 'E', 'R', 'T', 'Y', 'U', 'I', 'O', 'P']
NEW_LIST_SENTINEL = '+ New list…'
DEFAULT_LIST_NAME = 'My Pads'
DEFAULT_LIST_ID = 'list_default'

def _resource_dir():
    """Where bundled read-only assets (the icon) live. Under PyInstaller's
    onefile mode this is a temp extraction folder (sys._MEIPASS), not the
    .exe's own location."""
    if getattr(sys, 'frozen', False):
        return getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _data_dir():
    """Where user data (the config file) should persist — always next to
    the running .exe/.py, never the throwaway onefile extraction folder."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(_data_dir(), 'padboard_config.json')
ICON_PATH = os.path.join(_resource_dir(), 'songweaver_icon.ico')

BG = '#ffffff'
PANEL = '#f2f7f2'
PANEL_BORDER = '#e0e8e1'
PAD_BG = '#00a2e8'
PAD_BORDER = '#0072a6'
PAD_TEXT = '#ffffff'
PAD_TEXT_DIM = '#d0ecff'
ACCENT = '#16a34a'
ACCENT2 = '#4ade80'
PLAYING = '#14b8a6'          # teal — background of the currently playing/selected pad
PLAYING_BORDER = '#0f766e'   # darker teal border for the playing pad
SELECT_COLOR = '#f59e0b'     # amber ring for the keyboard-selected pad
TEXT = '#16281c'
TEXT_DIM = '#5c7263'
DANGER = '#dc2626'
SLIDER_BLUE = '#2563eb'
SLIDER_BLUE_ACTIVE = '#3b82f6'

PAD_W = 170
PAD_H = 130
CORNER_W_FRAC = 0.32
CORNER_H_FRAC = 0.24


class Tooltip:
    """A small hover label — Tkinter has no built-in tooltips."""

    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind('<Enter>', self._show, add='+')
        widget.bind('<Leave>', self._hide, add='+')
        widget.bind('<ButtonPress>', self._hide, add='+')

    def _show(self, _event=None):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 18
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f'+{x}+{y}')
        tk.Label(self.tip, text=self.text, bg='#222a26', fg='#ffffff',
                 font=('Segoe UI', 8), padx=8, pady=3).pack()

    def _hide(self, _event=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


def pump_com(duration=0.15):
    """Process pending COM/ActiveX messages for `duration` seconds.

    The Windows Media Player control is an out-of-process ActiveX object;
    without pumping, it gets stuck in the 'Transitioning' state and never
    reports Playing/Paused/error status. Tkinter's own event loop doesn't
    do this for us, so it's called after triggering playback and from the
    periodic poll loop.
    """
    start = time.time()
    while time.time() - start < duration:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.01)


def lighten(hex_color, factor):
    """Blend hex_color toward white by factor (0-1). Used for hover/gloss shades."""
    hex_color = hex_color.lstrip('#')
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return f'#{r:02x}{g:02x}{b:02x}'


def round_rect_points(x1, y1, x2, y2, r):
    return [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]


def probe_open(path, timeout=3.0):
    """Validate a file is actually playable, by really trying to play it
    briefly with a throwaway Windows Media Player instance.

    Returns (ok, error_message). WMP often reports no error at all for a
    file it can't decode (it just never leaves the 'Ready' state), so this
    doesn't trust the error queue alone — it waits for real confirmation
    that playback started.
    """
    wmp = win32com.client.Dispatch('WMPlayer.OCX')
    wmp.settings.autoStart = False
    try:
        wmp.error.clearErrorQueue()
    except Exception:
        pass
    try:
        wmp.URL = path
        wmp.controls.play()
    except Exception as e:
        return False, str(e)

    start = time.time()
    reached_playing = False
    while time.time() - start < timeout:
        pythoncom.PumpWaitingMessages()
        if wmp.playState == 3:  # Playing
            reached_playing = True
            break
        try:
            if wmp.error.errorCount > 0:
                return False, wmp.error.item(0).errorDescription
        except Exception:
            pass
        time.sleep(0.05)

    try:
        wmp.controls.stop()
        wmp.URL = ''
    except Exception:
        pass

    if reached_playing:
        return True, None
    return False, "Windows Media Player couldn't play this file (unsupported format or unreadable file)."


class PadSlot:
    """Metadata for one grid position in the currently displayed list.

    Holds no playback state — actual playback goes through a Player,
    independent of which list/pad slot is on screen (see PadBoardApp).
    """

    def __init__(self, index):
        self.index = index
        self.name = None
        self.path = None
        self.missing = False


class Player:
    """Wraps one Windows Media Player COM instance used for actual playback.

    Two instances are kept (crossfade in/out simultaneously) and reused
    across pads and lists, so a played track is entirely decoupled from
    which pad/list slot it came from.
    """

    def __init__(self):
        self._wmp = win32com.client.Dispatch('WMPlayer.OCX')
        self._wmp.settings.autoStart = False
        self.opened = False
        self.path = None

    def open(self, path):
        self.close()
        try:
            self._wmp.error.clearErrorQueue()
        except Exception:
            pass
        try:
            self._wmp.URL = path
        except Exception as e:
            self.opened = False
            return False, str(e)
        self.opened = True
        self.path = path
        return True, None

    def close(self):
        try:
            self._wmp.controls.stop()
            self._wmp.URL = ''
        except Exception:
            pass
        self.opened = False
        self.path = None

    def play(self):
        if not self.opened:
            return False
        try:
            self._wmp.controls.currentPosition = 0
            self._wmp.controls.play()
        except Exception:
            return False
        return True

    def pause(self):
        if self.opened:
            try:
                self._wmp.controls.pause()
            except Exception:
                pass

    def resume(self):
        if self.opened:
            try:
                self._wmp.controls.play()
            except Exception:
                pass

    def stop(self):
        if self.opened:
            try:
                self._wmp.controls.stop()
            except Exception:
                pass

    def _state(self):
        if not self.opened:
            return -1
        try:
            return self._wmp.playState
        except Exception:
            return -1

    def is_playing(self):
        return self._state() == 3

    def is_paused(self):
        return self._state() == 2

    def set_volume(self, vol_0_1000):
        if self.opened:
            try:
                self._wmp.settings.volume = max(0, min(100, int(vol_0_1000 / 10)))
            except Exception:
                pass

    def get_position(self):
        if not self.opened:
            return 0.0
        try:
            return float(self._wmp.controls.currentPosition)
        except Exception:
            return 0.0

    def get_duration(self):
        if not self.opened:
            return 0.0
        try:
            media = self._wmp.currentMedia
            return float(media.duration) if media else 0.0
        except Exception:
            return 0.0

    def seek(self, seconds):
        if self.opened:
            try:
                self._wmp.controls.currentPosition = max(0.0, seconds)
            except Exception:
                pass


class PadBoardApp:
    def __init__(self, root):
        self.root = root
        # All MAX_PAD_COUNT slots/widgets are pre-created once; a list only
        # shows/uses the first `current_pad_count` of them (see _apply_pad_count).
        self.pads = [PadSlot(i) for i in range(MAX_PAD_COUNT)]
        self.pad_widgets = []
        self.current_pad_count = BASE_PAD_COUNT
        self.master_volume = tk.IntVar(value=100)
        self.fade_in_sec = 2.0
        self.fade_out_sec = 2.0
        self.fade_jobs = {}  # id(player) -> root.after() job id
        self.all_lists_data = {}  # list_id -> {'name', 'pads': {index_str: {'name','path'}}, 'pad_count'}
        self.current_list_id = None
        self.lists = []  # [{'id':, 'name':}] mirror of all_lists_data, in display order

        # Playback is tracked independently of the displayed list/pads, so
        # switching lists never interrupts what's currently playing.
        self.player_a = Player()
        self.player_b = Player()
        self.now_playing = None  # {'list_id':, 'index':, 'name':, 'player':}
        self.seek_dragging = False
        self.play_all = False    # auto-play the next assigned pad when a song ends
        self.repeat_one = False  # replay the current song when it ends (wins over play_all)
        # Active custom play queue from the Play List dialog; drives song-to-song
        # advancement, overriding play_all/repeat: {'list_id':, 'order':[idx...], 'pos':}
        self.queue = None
        # Named, saved play lists per list: {list_id: [{'name':, 'order':[idx...]}]}
        self.saved_queues = {}
        self.selected_index = 0  # keyboard-selected pad (arrow-key cursor)

        self._build_ui()
        self._load_config()
        self._refresh_lists_array()
        self._refresh_list_combo()
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)
        self.root.bind_all('<Key>', self._on_key)
        # Safety net: a Combobox or Scale keeps keyboard focus after being
        # clicked/dragged, which silently swallows every shortcut in _on_key
        # from then on. Reclaim focus after any click that isn't genuine text
        # entry, deferred slightly so the click's own handling isn't disrupted.
        self.root.bind_all('<ButtonRelease-1>', self._reclaim_focus, add='+')
        self._update_mode_buttons()
        self._update_transport()
        self._render_selection()
        self._poll_playing()

    # ---------- UI ----------

    def _build_ui(self):
        self.root.title('SongWeaver - Local Soundboard')
        self.root.configure(bg=BG)
        self.root.geometry('940x830')
        self.root.minsize(780, 690)
        if os.path.exists(ICON_PATH):
            try:
                self.root.iconbitmap(ICON_PATH)
            except tk.TclError:
                pass

        header = tk.Frame(self.root, bg=BG)
        header.pack(fill='x', padx=24, pady=(18, 6))

        title = tk.Label(header, text='● SongWeaver', font=('Segoe UI', 16, 'bold'),
                          bg=BG, fg=TEXT)
        title.pack(side='left')

        settings_btn = tk.Button(header, text='⚙ Settings', command=self.open_settings,
                                  bg=PANEL, fg=TEXT, activebackground=PAD_BORDER, activeforeground=TEXT,
                                  relief='flat', padx=14, pady=7, font=('Segoe UI', 9))
        settings_btn.pack(side='right')

        queue_btn = tk.Button(header, text='▤ Play List', command=self.open_queue_modal,
                              bg=PANEL, fg=TEXT, activebackground=PAD_BORDER, activeforeground=TEXT,
                              relief='flat', padx=14, pady=7, font=('Segoe UI', 9))
        queue_btn.pack(side='right', padx=(0, 8))

        style = ttk.Style()
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass
        style.configure('PadBoard.TCombobox', fieldbackground=PANEL, background=PANEL, foreground=TEXT,
                         arrowcolor=TEXT, bordercolor=PANEL_BORDER, lightcolor=PANEL, darkcolor=PANEL,
                         selectbackground=PANEL, selectforeground=TEXT)

        self.list_combo = ttk.Combobox(header, state='readonly', width=22, style='PadBoard.TCombobox',
                                        font=('Segoe UI', 9))
        self.list_combo.pack(side='left', padx=(20, 6))
        self.list_combo.bind('<<ComboboxSelected>>', self._on_list_selected)

        rename_list_btn = tk.Button(header, text='✎', command=self.rename_list,
                                     bg=PANEL, fg=TEXT, activebackground=PAD_BORDER, activeforeground=TEXT,
                                     relief='flat', padx=10, pady=7, font=('Segoe UI', 9))
        rename_list_btn.pack(side='left')

        pad_count_bar = tk.Frame(self.root, bg=BG)
        pad_count_bar.pack(pady=(0, 8))

        self.pad_count_minus_btn = tk.Button(pad_count_bar, text='−', command=self.remove_pad_row,
                                              bg=PANEL, fg=TEXT, activebackground=PAD_BORDER, activeforeground=TEXT,
                                              relief='flat', width=2, font=('Segoe UI', 11, 'bold'),
                                              cursor='hand2')
        self.pad_count_minus_btn.pack(side='left', padx=(0, 10))

        self.pad_count_lbl = tk.Label(pad_count_bar, text=f'{BASE_PAD_COUNT} pads', font=('Segoe UI', 10, 'bold'),
                                       bg=BG, fg=TEXT, width=10)
        self.pad_count_lbl.pack(side='left')

        self.pad_count_plus_btn = tk.Button(pad_count_bar, text='+', command=self.add_pad_row,
                                             bg=PANEL, fg=TEXT, activebackground=PAD_BORDER, activeforeground=TEXT,
                                             relief='flat', width=2, font=('Segoe UI', 11, 'bold'),
                                             cursor='hand2')
        self.pad_count_plus_btn.pack(side='left', padx=(10, 0))

        grid_outer = tk.Frame(self.root, bg=BG)
        grid_outer.pack(fill='both', expand=True, padx=24, pady=10)

        self.grid_canvas = tk.Canvas(grid_outer, bg=BG, highlightthickness=0)
        grid_scroll = tk.Scrollbar(grid_outer, orient='vertical', command=self.grid_canvas.yview)
        self.grid_canvas.configure(yscrollcommand=grid_scroll.set)
        self.grid_canvas.pack(side='left', fill='both', expand=True)
        grid_scroll.pack(side='right', fill='y')

        grid_frame = tk.Frame(self.grid_canvas, bg=BG)
        grid_window = self.grid_canvas.create_window((0, 0), window=grid_frame, anchor='nw')
        grid_frame.bind('<Configure>', lambda e: self.grid_canvas.configure(
            scrollregion=self.grid_canvas.bbox('all')))
        self.grid_canvas.bind('<Configure>', lambda e: self.grid_canvas.itemconfig(grid_window, width=e.width))

        def _on_mousewheel(event):
            self.grid_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        self.grid_canvas.bind('<Enter>', lambda e: self.grid_canvas.bind_all('<MouseWheel>', _on_mousewheel))
        self.grid_canvas.bind('<Leave>', lambda e: self.grid_canvas.unbind_all('<MouseWheel>'))

        for i in range(MAX_PAD_COUNT):
            pad_widget = self._build_pad(grid_frame, i)
            self.pad_widgets.append(pad_widget)

        footer = tk.Label(
            self.root,
            text=('Click an empty pad to assign a song  ·  Click a loaded pad to play/stop  ·  '
                  'Double-click the name to rename  ·  ✕ clears a pad\n'
                  'Keys 1-0 and Q-P trigger pads 1-20. Only one song plays at a time - starting a new one '
                  'crossfades out of the current one, even across a list switch (adjust timing in ⚙ Settings).\n'
                  'Use the list dropdown to switch pad sets, ✎ to rename the current one, or '
                  '"+ New list…" for a fresh set of 20 pads. Use − / + above the grid to remove or add a row '
                  f'of {PAD_INCREMENT}, from {BASE_PAD_COUNT} up to {MAX_PAD_COUNT} pads per list.'),
            font=('Segoe UI', 8), bg=BG, fg=TEXT_DIM, justify='center'
        )
        footer.pack(pady=(4, 10))

        self._build_transport()

    def _build_transport(self):
        bar = tk.Frame(self.root, bg=PANEL, highlightthickness=1, highlightbackground=PANEL_BORDER)
        bar.pack(side='bottom', fill='x')

        inner = tk.Frame(bar, bg=PANEL)
        inner.pack(fill='x', padx=22, pady=12)

        # Now playing (left)
        now_frame = tk.Frame(inner, bg=PANEL)
        now_frame.pack(side='left')

        self.now_icon_lbl = tk.Label(now_frame, text='♪', font=('Segoe UI', 13), bg=PAD_BG, fg=PAD_TEXT,
                                      width=2, height=1)
        self.now_icon_lbl.pack(side='left', padx=(0, 10))

        now_text = tk.Frame(now_frame, bg=PANEL)
        now_text.pack(side='left')
        self.now_title_lbl = tk.Label(now_text, text='No song selected', font=('Segoe UI', 10, 'bold'),
                                       bg=PANEL, fg=TEXT, anchor='w', width=18)
        self.now_title_lbl.pack(anchor='w')
        self.now_sub_lbl = tk.Label(now_text, text='Pick a pad to get started', font=('Segoe UI', 8),
                                     bg=PANEL, fg=TEXT_DIM, anchor='w', width=20)
        self.now_sub_lbl.pack(anchor='w')

        # Transport buttons + seek bar (center, expands to fill remaining space)
        btns = tk.Frame(inner, bg=PANEL)
        btns.pack(side='left', fill='x', expand=True, padx=20)

        self.play_all_btn = tk.Button(btns, text='⏩', command=self.toggle_play_all,
                                       bg=PANEL, fg=TEXT_DIM, activebackground=PANEL_BORDER,
                                       relief='flat', bd=1, width=2, height=1, font=('Segoe UI', 12),
                                       cursor='hand2')
        self.play_all_btn.pack(side='left', padx=(0, 6))

        self.repeat_btn = tk.Button(btns, text='🔂', command=self.toggle_repeat,
                                     bg=PANEL, fg=TEXT_DIM, activebackground=PANEL_BORDER,
                                     relief='flat', bd=1, width=2, height=1, font=('Segoe UI', 12),
                                     cursor='hand2')
        self.repeat_btn.pack(side='left', padx=(0, 12))

        self.play_pause_btn = tk.Button(btns, text='▶', command=self.on_transport_play_pause,
                                         bg=ACCENT, fg='#ffffff', activebackground=ACCENT2,
                                         activeforeground='#ffffff', relief='flat', bd=0,
                                         width=3, height=1, font=('Segoe UI', 18, 'bold'),
                                         disabledforeground='#ffffff', state='disabled', cursor='hand2')
        self.play_pause_btn.pack(side='left', padx=(0, 10))

        self.next_btn = tk.Button(btns, text='⏭', command=self.on_transport_next,
                                   bg=BG, fg=TEXT, activebackground=PANEL_BORDER, activeforeground=TEXT,
                                   relief='flat', bd=1, width=3, height=1, font=('Segoe UI', 14),
                                   state='disabled', cursor='hand2')
        self.next_btn.pack(side='left', padx=(0, 10))

        Tooltip(self.play_all_btn, 'Play All — auto-play the next pad when a song ends')
        Tooltip(self.repeat_btn, 'Repeat — replay the current song')
        Tooltip(self.play_pause_btn, 'Play / Pause')
        Tooltip(self.next_btn, 'Play Next — skip to the next song')

        seek_frame = tk.Frame(btns, bg=PANEL)
        seek_frame.pack(side='left', fill='x', expand=True, padx=(0, 10))

        self.seek_current_lbl = tk.Label(seek_frame, text='0:00', font=('Consolas', 8), bg=PANEL, fg=TEXT_DIM,
                                          width=4)
        self.seek_current_lbl.pack(side='left')

        self.seek_var = tk.DoubleVar(value=0.0)
        self.seek_scale = tk.Scale(seek_frame, from_=0, to=100, orient='horizontal', variable=self.seek_var,
                                    bg=PANEL, fg=SLIDER_BLUE, troughcolor='#dfe8e0', highlightthickness=0,
                                    sliderlength=14, showvalue=False, bd=0, resolution=0.1,
                                    activebackground=SLIDER_BLUE_ACTIVE, state='disabled',
                                    command=self._on_seek_drag)
        self.seek_scale.pack(side='left', fill='x', expand=True, padx=6)
        self.seek_scale.bind('<ButtonPress-1>', self._on_seek_press)
        self.seek_scale.bind('<ButtonRelease-1>', self._on_seek_release)
        # A Scale keeps keyboard focus after being dragged, which would otherwise
        # silently swallow every keyboard shortcut from here on (see _on_key).
        self.seek_scale.bind('<ButtonRelease-1>', lambda e: self.root.focus_set(), add='+')

        self.seek_duration_lbl = tk.Label(seek_frame, text='0:00', font=('Consolas', 8), bg=PANEL, fg=TEXT_DIM,
                                           width=4)
        self.seek_duration_lbl.pack(side='left')

        self.stop_btn = tk.Button(btns, text='⏹', command=self.on_transport_stop,
                                   bg=BG, fg=TEXT, activebackground=PANEL_BORDER, activeforeground=TEXT,
                                   relief='flat', bd=1, width=3, height=1, font=('Segoe UI', 15),
                                   state='disabled', cursor='hand2')
        self.stop_btn.pack(side='left')
        Tooltip(self.stop_btn, 'Stop')

        # Volume (right)
        vol_frame = tk.Frame(inner, bg=PANEL)
        vol_frame.pack(side='right')

        tk.Label(vol_frame, text='🔊', font=('Segoe UI', 12), bg=PANEL, fg=TEXT_DIM).pack(side='left', padx=(0, 8))
        vol_scale = tk.Scale(vol_frame, from_=0, to=100, orient='horizontal', variable=self.master_volume,
                              bg=PANEL, fg=SLIDER_BLUE, troughcolor='#dfe8e0', highlightthickness=0,
                              length=150, sliderlength=18, showvalue=False, bd=0,
                              activebackground=SLIDER_BLUE_ACTIVE, command=self._on_volume_change)
        vol_scale.pack(side='left')
        vol_scale.bind('<ButtonRelease-1>', lambda e: self.root.focus_set())
        self.vol_pct_lbl = tk.Label(vol_frame, text='100%', font=('Consolas', 9), bg=PANEL, fg=TEXT_DIM, width=4)
        self.vol_pct_lbl.pack(side='left', padx=(8, 0))

    def _build_pad(self, parent, index):
        w, h, r = PAD_W, PAD_H, 16
        canvas = tk.Canvas(parent, width=w, height=h, bg=BG, highlightthickness=0, cursor='hand2')

        bg_id = canvas.create_polygon(round_rect_points(2, 2, w - 2, h - 2, r), smooth=True,
                                       fill=PAD_BG, outline=PAD_BORDER, width=2, tags=('padclick',))
        gloss_id = canvas.create_oval(12, 8, w * 0.62, h * 0.38, fill=lighten(PAD_BG, 0.22),
                                       outline='', tags=('padclick',))

        key_id = canvas.create_text(16, 16, text=(KEYS[index] if index < len(KEYS) else ''), fill=PAD_TEXT_DIM,
                                     font=('Consolas', 9), anchor='w', tags=('padclick',))
        num_id = canvas.create_text(w - 14, 16, text=str(index + 1), fill=PAD_TEXT_DIM,
                                     font=('Segoe UI', 9), anchor='e', tags=('padclick',))

        icon_bg_id = canvas.create_oval(w / 2 - 24, 34, w / 2 + 24, 82,
                                         fill=lighten(PAD_BG, 0.14), outline='', tags=('padclick',))
        icon_id = canvas.create_text(w / 2, 58, text='+', fill=PAD_TEXT_DIM,
                                      font=('Segoe UI', 20), tags=('padclick',))

        name_id = canvas.create_text(w / 2, 100, text='Add song', fill=PAD_TEXT_DIM,
                                      font=('Segoe UI', 12), width=w - 22,
                                      justify='center', tags=('padclick',))

        clear_id = canvas.create_text(w - 14, h - 12, text='', fill=PAD_TEXT_DIM,
                                       font=('Segoe UI', 9), tags=('clearbtn',))

        # amber ring shown when this pad is the keyboard-selected one (hidden by default)
        sel_id = canvas.create_polygon(round_rect_points(3, 3, w - 3, h - 3, r - 1), smooth=True,
                                        fill='', outline=SELECT_COLOR, width=3, state='hidden')

        canvas.tag_bind('padclick', '<Button-1>', lambda e, i=index: self._on_pad_canvas_click(e, i))
        canvas.tag_bind('clearbtn', '<Button-1>', lambda e, i=index: self.clear_pad(i))
        canvas.tag_bind(name_id, '<Double-Button-1>', lambda e, i=index: self.rename_pad(i))
        canvas.bind('<Enter>', lambda e, i=index: self._pad_hover(i, True))
        canvas.bind('<Leave>', lambda e, i=index: self._pad_hover(i, False))

        return {
            'canvas': canvas, 'bg': bg_id, 'gloss': gloss_id, 'key': key_id, 'num': num_id,
            'icon_bg': icon_bg_id, 'icon': icon_id, 'name': name_id, 'clear': clear_id, 'sel': sel_id,
        }

    def _pad_hover(self, index, entering):
        w = self.pad_widgets[index]
        base = PLAYING if self._is_now_playing(index) else PAD_BG
        w['canvas'].itemconfig(w['bg'], fill=(lighten(base, 0.08) if entering else base))

    # ---------- keyboard-selected pad (arrow-key cursor) ----------

    def _render_selection(self):
        for i, w in enumerate(self.pad_widgets):
            state = 'normal' if (i == self.selected_index and i < self.current_pad_count) else 'hidden'
            w['canvas'].itemconfig(w['sel'], state=state)

    def _set_selected(self, index):
        if index < 0 or index >= self.current_pad_count:
            return
        self.selected_index = index
        self._render_selection()
        self._scroll_to_selected()

    def _move_selection(self, delta):
        target = self.selected_index + delta
        target = max(0, min(self.current_pad_count - 1, target))
        self._set_selected(target)

    def _scroll_to_selected(self):
        try:
            canvas = self.grid_canvas
            widget = self.pad_widgets[self.selected_index]['canvas']
            self.root.update_idletasks()
            total = widget.master.winfo_height()
            if total <= 1:
                return
            top = widget.winfo_y() / total
            bottom = (widget.winfo_y() + widget.winfo_height()) / total
            view_lo, view_hi = canvas.yview()
            if top < view_lo:
                canvas.yview_moveto(max(0.0, top - 0.02))
            elif bottom > view_hi:
                canvas.yview_moveto(min(1.0, bottom - (view_hi - view_lo) + 0.02))
        except (tk.TclError, ZeroDivisionError):
            pass

    def _on_pad_canvas_click(self, event, index):
        # Reclaim keyboard focus from whatever had it (list dropdown, a slider)
        # so shortcuts keep working right after interacting with a pad — see
        # the focus guard in _on_key.
        self.root.focus_set()
        # Top-left/top-right corners (where the key/number badges sit) are a
        # dead zone — assigning or playing only triggers from the pad's body.
        corner_w = PAD_W * CORNER_W_FRAC
        corner_h = PAD_H * CORNER_H_FRAC
        in_top_left = event.x <= corner_w and event.y <= corner_h
        in_top_right = event.x >= (PAD_W - corner_w) and event.y <= corner_h
        if in_top_left or in_top_right:
            return
        self.on_pad_click(index)

    def _apply_pad_count(self):
        count = self.current_pad_count
        for i, w in enumerate(self.pad_widgets):
            r, c = divmod(i, COLS)
            if i < count:
                w['canvas'].grid(row=r, column=c, padx=8, pady=8)
            else:
                w['canvas'].grid_remove()

        self.pad_count_lbl.config(text=f'{count} pads')
        self.pad_count_minus_btn.config(state=('disabled' if count <= BASE_PAD_COUNT else 'normal'))
        self.pad_count_plus_btn.config(state=('disabled' if count >= MAX_PAD_COUNT else 'normal'))

    def add_pad_row(self):
        entry = self.all_lists_data.get(self.current_list_id)
        if not entry:
            return
        current = entry.get('pad_count', BASE_PAD_COUNT)
        if current >= MAX_PAD_COUNT:
            return
        entry['pad_count'] = min(MAX_PAD_COUNT, current + PAD_INCREMENT)
        self._save_config()
        self._load_pads_for_current_list()

    def remove_pad_row(self):
        entry = self.all_lists_data.get(self.current_list_id)
        if not entry:
            return
        current = entry.get('pad_count', BASE_PAD_COUNT)
        if current <= BASE_PAD_COUNT:
            return
        # Just hides the last row — pad data underneath it is kept, so adding
        # the row back later restores whatever was assigned there.
        entry['pad_count'] = max(BASE_PAD_COUNT, current - PAD_INCREMENT)
        self._save_config()
        self._load_pads_for_current_list()

    def _is_now_playing(self, index):
        return (self.now_playing is not None and self.now_playing['list_id'] == self.current_list_id
                and self.now_playing['index'] == index)

    def _render_pad(self, index):
        pad = self.pads[index]
        w = self.pad_widgets[index]
        c = w['canvas']
        is_current = self._is_now_playing(index)
        playing = is_current and self.now_playing['player'].is_playing()

        # The currently playing/selected pad turns teal.
        base = PLAYING if is_current else PAD_BG
        c.itemconfig(w['bg'], fill=base)
        c.itemconfig(w['gloss'], fill=lighten(base, 0.22))
        c.itemconfig(w['icon_bg'], fill=lighten(base, 0.14))

        if pad.name:
            c.itemconfig(w['icon'], text=('⏸' if playing else '▶'), fill=PAD_TEXT)
            display_name = pad.name if not pad.missing else f'{pad.name} (missing)'
            c.itemconfig(w['name'], text=display_name, fill=(DANGER if pad.missing else PAD_TEXT))
            c.itemconfig(w['clear'], text='✕')
            border = PLAYING_BORDER if is_current else PAD_BORDER
            c.itemconfig(w['bg'], outline=border, width=(3 if is_current else 2))
        else:
            c.itemconfig(w['icon'], text='+', fill=PAD_TEXT_DIM)
            c.itemconfig(w['name'], text='Add song', fill=PAD_TEXT_DIM)
            c.itemconfig(w['clear'], text='')
            c.itemconfig(w['bg'], outline=PAD_BORDER, width=2)

    def _update_transport(self):
        record = self.now_playing
        if not record:
            self.now_icon_lbl.config(text='♪')
            self.now_title_lbl.config(text='No song selected')
            self.now_sub_lbl.config(text='Pick a pad to get started')
            self.play_pause_btn.config(text='▶', state='disabled')
            self.stop_btn.config(state='disabled')
            self.next_btn.config(state='disabled')
            self._update_seek()
            return
        paused = record['player'].is_paused()
        list_name = self.all_lists_data.get(record['list_id'], {}).get('name', '')
        prefix = f'{list_name}  ·  ' if list_name else ''
        queue_tag = ''
        if self.queue and self.queue['list_id'] == record['list_id']:
            queue_tag = f"  ·  queue {self.queue['pos'] + 1}/{len(self.queue['order'])}"
        self.now_icon_lbl.config(text=('▶' if paused else '⏸'))
        self.now_title_lbl.config(text=record['name'])
        self.now_sub_lbl.config(text=f"{prefix}Pad {record['index'] + 1}  ·  "
                                     + ('paused' if paused else 'playing') + queue_tag)
        self.play_pause_btn.config(text=('▶' if paused else '⏸'), state='normal')
        self.stop_btn.config(state='normal')
        self.next_btn.config(state='normal')
        self._update_seek()

    def _update_seek(self):
        record = self.now_playing
        player = record['player'] if record else None
        duration = player.get_duration() if player else 0.0
        if not player or duration <= 0:
            self.seek_scale.config(to=100, state='disabled')
            if not self.seek_dragging:
                self.seek_var.set(0)
            self.seek_current_lbl.config(text='0:00')
            self.seek_duration_lbl.config(text='0:00')
            return
        self.seek_scale.config(to=duration, state='normal')
        if not self.seek_dragging:
            position = player.get_position()
            self.seek_var.set(position)
            self.seek_current_lbl.config(text=self._format_time(position))
        self.seek_duration_lbl.config(text=self._format_time(duration))

    def _on_seek_drag(self, value):
        if self.seek_dragging:
            try:
                self.seek_current_lbl.config(text=self._format_time(float(value)))
            except (TypeError, ValueError):
                pass

    def _on_seek_press(self, _event):
        self.seek_dragging = True

    def _on_seek_release(self, _event):
        self.seek_dragging = False
        record = self.now_playing
        if record:
            record['player'].seek(self.seek_var.get())

    @staticmethod
    def _format_time(seconds):
        try:
            seconds = max(0, int(seconds))
        except (TypeError, ValueError):
            seconds = 0
        return f'{seconds // 60}:{seconds % 60:02d}'

    # ---------- play modes ----------

    def toggle_play_all(self):
        self.play_all = not self.play_all
        self._update_mode_buttons()
        self._save_config()

    def toggle_repeat(self):
        self.repeat_one = not self.repeat_one
        self._update_mode_buttons()
        self._save_config()

    def _update_mode_buttons(self):
        for btn, on in ((self.play_all_btn, self.play_all), (self.repeat_btn, self.repeat_one)):
            if on:
                btn.config(bg=SLIDER_BLUE, fg='#ffffff', activebackground=SLIDER_BLUE_ACTIVE)
            else:
                btn.config(bg=PANEL, fg=TEXT_DIM, activebackground=PANEL_BORDER)

    def _next_assigned_pad(self, list_id, after_index):
        """Return (index, name, path) of the next assigned+existing pad after
        `after_index` within a list, or None if there isn't one."""
        entry = self.all_lists_data.get(list_id)
        if not entry:
            return None
        pads_data = entry.get('pads', {})
        pad_count = entry.get('pad_count', BASE_PAD_COUNT)
        for i in range(after_index + 1, pad_count):
            info = pads_data.get(str(i))
            if info and info.get('path') and os.path.exists(info['path']):
                return (i, info.get('name'), info['path'])
        return None

    # ---------- play queue (Play List dialog) ----------

    def _play_queue_item(self, list_id, index, crossfade=False):
        entry = self.all_lists_data.get(list_id)
        info = entry.get('pads', {}).get(str(index)) if entry else None
        if info and info.get('path') and os.path.exists(info['path']):
            self._start_track(list_id, index, info.get('name'), info['path'], crossfade=crossfade)
        else:
            self._advance_queue()  # queued pad is gone — skip to the next one

    def _advance_queue(self):
        if not self.queue:
            return
        self.queue['pos'] += 1
        if self.queue['pos'] < len(self.queue['order']):
            self._play_queue_item(self.queue['list_id'], self.queue['order'][self.queue['pos']])
        else:
            self.queue = None  # reached the end of the queue

    def open_queue_modal(self):
        entry = self.all_lists_data.get(self.current_list_id)
        if not entry:
            return
        pads_data = entry.get('pads', {})
        pad_count = entry.get('pad_count', BASE_PAD_COUNT)

        def assigned(i):
            info = pads_data.get(str(i))
            return info.get('name') if info and info.get('path') else None

        # Prefill the working order from an active queue for this list, if any.
        working = []
        if self.queue and self.queue['list_id'] == self.current_list_id:
            working = [i for i in self.queue['order'] if assigned(i)]

        win = tk.Toplevel(self.root)
        win.title('Play List')
        win.configure(bg=PANEL)
        win.transient(self.root)
        win.resizable(False, False)
        if os.path.exists(ICON_PATH):
            try:
                win.iconbitmap(ICON_PATH)
            except tk.TclError:
                pass

        tk.Label(win, text='Play List', font=('Segoe UI', 13, 'bold'), bg=PANEL, fg=TEXT).pack(
            padx=18, pady=(16, 2), anchor='w')
        tk.Label(win, text='Add songs (double-click to add fast), arrange in order, then Play. '
                 'Save a queue to reuse it later.',
                 font=('Segoe UI', 8), bg=PANEL, fg=TEXT_DIM).pack(padx=18, anchor='w')

        saved_row = tk.Frame(win, bg=PANEL)
        saved_row.pack(fill='x', padx=18, pady=(10, 0))
        tk.Label(saved_row, text='Saved lists:', font=('Segoe UI', 9, 'bold'),
                 bg=PANEL, fg=TEXT_DIM).pack(side='left')
        saved_var = tk.StringVar()
        saved_combo = ttk.Combobox(saved_row, textvariable=saved_var, state='readonly', width=20,
                                   style='PadBoard.TCombobox', font=('Segoe UI', 9))
        saved_combo.pack(side='left', padx=(6, 6))

        body = tk.Frame(win, bg=PANEL)
        body.pack(padx=18, pady=12)

        lb_opts = dict(width=30, height=14, activestyle='none', exportselection=False,
                       bg=BG, fg=TEXT, highlightthickness=1, highlightbackground=PANEL_BORDER,
                       selectbackground=ACCENT, selectforeground='#ffffff', relief='flat',
                       font=('Segoe UI', 9))

        left_frame = tk.Frame(body, bg=PANEL)
        left_frame.grid(row=0, column=0, sticky='n')
        tk.Label(left_frame, text='All songs', font=('Segoe UI', 9, 'bold'), bg=PANEL, fg=TEXT_DIM).pack(anchor='w')
        left_lb = tk.Listbox(left_frame, **lb_opts)
        left_lb.pack()

        mid = tk.Frame(body, bg=PANEL)
        mid.grid(row=0, column=1, padx=12)

        right_frame = tk.Frame(body, bg=PANEL)
        right_frame.grid(row=0, column=2, sticky='n')
        tk.Label(right_frame, text='Play queue (in order)', font=('Segoe UI', 9, 'bold'),
                 bg=PANEL, fg=TEXT_DIM).pack(anchor='w')
        right_lb = tk.Listbox(right_frame, **lb_opts)
        right_lb.pack()

        left_map = []  # pad index for each row in the left listbox

        def refresh():
            left_lb.delete(0, 'end')
            left_map.clear()
            for i in range(pad_count):
                nm = assigned(i)
                if nm and i not in working:
                    left_map.append(i)
                    left_lb.insert('end', f'{i + 1}. {nm}')
            right_lb.delete(0, 'end')
            for pos, i in enumerate(working):
                nm = assigned(i) or f'Pad {i + 1}'
                right_lb.insert('end', f'{pos + 1}. {nm}')

        def do_add():
            sel = left_lb.curselection()
            if not sel:
                return
            idx = left_map[sel[0]]
            if idx not in working:
                working.append(idx)
            refresh()

        def do_remove():
            sel = right_lb.curselection()
            if not sel:
                return
            working.pop(sel[0])
            refresh()

        def do_move(delta):
            sel = right_lb.curselection()
            if not sel:
                return
            p = sel[0]
            t = p + delta
            if t < 0 or t >= len(working):
                return
            working[p], working[t] = working[t], working[p]
            refresh()
            right_lb.selection_set(t)

        btn_opts = dict(bg=PANEL, fg=TEXT, activebackground=PAD_BORDER, relief='flat',
                        font=('Segoe UI', 12), width=3, cursor='hand2',
                        highlightthickness=1, highlightbackground=PANEL_BORDER)
        tk.Frame(mid, bg=PANEL, height=20).pack()
        tk.Button(mid, text='→', command=do_add, **btn_opts).pack(pady=6)
        tk.Button(mid, text='←', command=do_remove, **btn_opts).pack(pady=6)

        left_lb.bind('<Double-Button-1>', lambda e: do_add())
        right_lb.bind('<Double-Button-1>', lambda e: do_remove())

        order_frame = tk.Frame(right_frame, bg=PANEL)
        order_frame.pack(fill='x', pady=(6, 0))
        ord_opts = dict(bg=PANEL, fg=TEXT, activebackground=PAD_BORDER, relief='flat',
                        font=('Segoe UI', 9), cursor='hand2',
                        highlightthickness=1, highlightbackground=PANEL_BORDER)
        tk.Button(order_frame, text='▲ Up', command=lambda: do_move(-1), **ord_opts).pack(side='left', expand=True, fill='x', padx=(0, 3))
        tk.Button(order_frame, text='▼ Down', command=lambda: do_move(1), **ord_opts).pack(side='left', expand=True, fill='x', padx=(3, 0))

        actions = tk.Frame(win, bg=PANEL)
        actions.pack(fill='x', padx=18, pady=(4, 16))

        def do_ok():
            win.destroy()
            if not working:
                self.queue = None
                return
            self.queue = {'list_id': self.current_list_id, 'order': list(working), 'pos': 0}
            self._play_queue_item(self.current_list_id, working[0], crossfade=True)

        def do_apply():
            # Save/apply the list without interrupting the current song.
            win.destroy()
            if not working:
                self.queue = None
                self._update_transport()
                return
            order = list(working)
            np = self.now_playing
            if np and np['list_id'] == self.current_list_id:
                pos = order.index(np['index']) if np['index'] in order else 0
                self.queue = {'list_id': self.current_list_id, 'order': order, 'pos': pos}
                self._update_transport()
            elif np:
                self.queue = {'list_id': self.current_list_id, 'order': order, 'pos': 0}
                self._update_transport()
            else:
                self.queue = {'list_id': self.current_list_id, 'order': order, 'pos': 0}
                self._play_queue_item(self.current_list_id, order[0], crossfade=True)

        # ---- saved play lists ----
        new_label = '— New list —'

        def current_saved():
            return self.saved_queues.get(self.current_list_id, [])

        def refresh_saved():
            saved_combo['values'] = [new_label] + [q['name'] for q in current_saved()]
            saved_combo.current(0)

        def on_saved_select(_e=None):
            name = saved_var.get()
            if not name or name == new_label:
                return
            q = next((x for x in current_saved() if x['name'] == name), None)
            if q:
                working.clear()
                working.extend(i for i in q['order'] if assigned(i))
                refresh()

        def do_save():
            if not working:
                messagebox.showinfo('SongWeaver', 'Add some songs to the queue first.', parent=win)
                return
            suggested = saved_var.get() if saved_var.get() != new_label else f'List {len(current_saved()) + 1}'
            name = simpledialog.askstring('Save play list', 'Save this play list as:',
                                          initialvalue=suggested, parent=win)
            if not name or not name.strip():
                return
            trimmed = name.strip()
            arr = self.saved_queues.setdefault(self.current_list_id, [])
            existing = next((x for x in arr if x['name'] == trimmed), None)
            if existing:
                existing['order'] = list(working)
            else:
                arr.append({'name': trimmed, 'order': list(working)})
            self._save_config()
            refresh_saved()
            saved_var.set(trimmed)

        def do_delete():
            name = saved_var.get()
            if not name or name == new_label:
                return
            if messagebox.askyesno('Delete', f'Delete saved play list "{name}"?', parent=win):
                self.saved_queues[self.current_list_id] = [x for x in current_saved() if x['name'] != name]
                self._save_config()
                refresh_saved()

        saved_combo.bind('<<ComboboxSelected>>', on_saved_select)
        saved_btn_opts = dict(bg=PANEL, fg=TEXT, activebackground=PAD_BORDER, relief='flat',
                              padx=10, pady=4, font=('Segoe UI', 9), cursor='hand2',
                              highlightthickness=1, highlightbackground=PANEL_BORDER)
        tk.Button(saved_row, text='💾 Save', command=do_save, **saved_btn_opts).pack(side='left')
        tk.Button(saved_row, text='🗑 Delete', command=do_delete, **saved_btn_opts).pack(side='left', padx=(6, 0))

        tk.Button(actions, text='▶ Play now', command=do_ok, bg=ACCENT, fg='#ffffff',
                  activebackground=ACCENT2, activeforeground='#ffffff', relief='flat',
                  padx=16, pady=6, font=('Segoe UI', 9, 'bold')).pack(side='right')
        ok_btn = tk.Button(actions, text='OK', command=do_apply, bg=PANEL, fg=TEXT,
                           activebackground=PAD_BORDER, relief='flat', padx=16, pady=6,
                           font=('Segoe UI', 9), highlightthickness=1, highlightbackground=PANEL_BORDER)
        ok_btn.pack(side='right', padx=(0, 8))
        Tooltip(ok_btn, 'Save the list without interrupting the song playing now')
        tk.Button(actions, text='Cancel', command=win.destroy, bg=PANEL, fg=TEXT,
                  activebackground=PAD_BORDER, relief='flat', padx=16, pady=6,
                  font=('Segoe UI', 9), highlightthickness=1, highlightbackground=PANEL_BORDER).pack(side='right', padx=(0, 8))

        refresh_saved()
        refresh()
        win.protocol('WM_DELETE_WINDOW', win.destroy)

    # ---------- actions ----------

    def on_pad_click(self, index):
        pad = self.pads[index]
        if pad.name and not pad.missing:
            self.activate_pad(index)
        else:
            self.assign_pad(index)

    def assign_pad(self, index):
        path = filedialog.askopenfilename(
            title='Choose a song',
            filetypes=[('Audio files', '*.mp3 *.wav *.wma *.m4a'), ('All files', '*.*')]
        )
        if not path:
            return
        ok, err_msg = probe_open(path)
        if not ok:
            messagebox.showerror('SongWeaver', f"Couldn't load this file:\n{os.path.basename(path)}\n\n{err_msg}\n\n"
                                              "The file may be corrupt, DRM-protected, or missing the codec "
                                              "needed to play it.")
            return
        pad = self.pads[index]
        pad.name = os.path.splitext(os.path.basename(path))[0]
        pad.path = path
        pad.missing = False
        self._save_config()
        self._render_pad(index)

    def clear_pad(self, index):
        if self._is_now_playing(index):
            self._fade_out_and_stop_now_playing()
        pad = self.pads[index]
        pad.name = None
        pad.path = None
        pad.missing = False
        self._save_config()
        self._render_pad(index)

    def rename_pad(self, index):
        pad = self.pads[index]
        if not pad.name:
            return
        new_name = simpledialog.askstring('Rename pad', 'Pad name:', initialvalue=pad.name, parent=self.root)
        if new_name:
            pad.name = new_name.strip() or pad.name
            self._save_config()
            self._render_pad(index)
            if self._is_now_playing(index):
                self.now_playing['name'] = pad.name
                self._update_transport()

    def activate_pad(self, index):
        """Only one track plays at a time, tracked independently of the
        displayed list so switching lists never interrupts it. Starting a
        pad crossfades out whatever's currently playing while fading the
        new one in."""
        pad = self.pads[index]
        if not pad.name or pad.missing:
            return

        # A manual pad tap ends any running play-queue.
        self.queue = None

        if self._is_now_playing(index):
            self._fade_out_and_stop_now_playing()
            return

        self._start_track(self.current_list_id, index, pad.name, pad.path, crossfade=True)

    def _start_track(self, list_id, index, name, path, crossfade=True):
        """Play a specific track. Used both for manual pad clicks (crossfade
        out of the current song) and for Play All auto-advance (no crossfade —
        the previous song already ended)."""
        previous = self.now_playing
        incoming = self.player_b if (previous and previous['player'] is self.player_a) else self.player_a

        ok, err_msg = incoming.open(path)
        if not ok:
            messagebox.showerror('SongWeaver', f"Couldn't play this file:\n{name}\n\n{err_msg}")
            return False

        incoming.set_volume(0)
        incoming.play()
        pump_com(0.15)
        self.now_playing = {'list_id': list_id, 'index': index, 'name': name,
                             'player': incoming, 'started_at': time.time()}
        if crossfade and previous is not None:
            self._fade_out_player_record(previous)
        self._fade_in_player(incoming)
        if list_id == self.current_list_id:
            self._render_pad(index)
            self._set_selected(index)  # the playing pad becomes the selected one
        self._update_transport()
        return True

    def _fade_in_player(self, player):
        self._cancel_fade(player)
        target = self.master_volume.get() / 100.0
        self._fade(player, 0.0, target, self.fade_in_sec)

    def _fade_out_player_record(self, record):
        player = record['player']
        self._cancel_fade(player)
        start_vol = self.master_volume.get() / 100.0

        def done():
            player.stop()
            player.close()
            if self.now_playing is record:
                self.now_playing = None
            if record['list_id'] == self.current_list_id:
                self._render_pad(record['index'])
            self._update_transport()

        self._fade(player, start_vol, 0.0, self.fade_out_sec, on_complete=done)

    def _fade_out_and_stop_now_playing(self):
        if not self.now_playing:
            return
        self._fade_out_player_record(self.now_playing)
        self._update_transport()

    def _fade(self, player, start_vol, end_vol, duration, on_complete=None):
        key = id(player)
        if duration <= 0:
            player.set_volume(end_vol * 1000)
            if on_complete:
                on_complete()
            return
        steps = max(1, round(duration * 25))
        step_ms = max(10, int(duration * 1000 / steps))
        state = {'i': 0}

        def step():
            state['i'] += 1
            t = state['i'] / steps
            vol = start_vol + (end_vol - start_vol) * t
            player.set_volume(vol * 1000)
            if state['i'] >= steps:
                player.set_volume(end_vol * 1000)
                self.fade_jobs.pop(key, None)
                if on_complete:
                    on_complete()
            else:
                self.fade_jobs[key] = self.root.after(step_ms, step)

        step()

    def _cancel_fade(self, player):
        key = id(player)
        job = self.fade_jobs.pop(key, None)
        if job:
            try:
                self.root.after_cancel(job)
            except (tk.TclError, ValueError):
                pass

    def open_settings(self):
        win = tk.Toplevel(self.root)
        win.title('Fade Settings')
        win.configure(bg=PANEL)
        win.resizable(False, False)
        win.transient(self.root)

        fade_in_var = tk.DoubleVar(value=self.fade_in_sec)
        fade_out_var = tk.DoubleVar(value=self.fade_out_sec)

        tk.Label(win, text='Fade In (seconds)', bg=PANEL, fg=TEXT, font=('Segoe UI', 9)).pack(
            padx=18, pady=(18, 2), anchor='w')
        tk.Scale(win, from_=0, to=10, resolution=0.5, orient='horizontal', variable=fade_in_var,
                 bg=PANEL, fg=ACCENT, troughcolor=BG, highlightthickness=0, length=220).pack(padx=18)

        tk.Label(win, text='Fade Out (seconds)', bg=PANEL, fg=TEXT, font=('Segoe UI', 9)).pack(
            padx=18, pady=(14, 2), anchor='w')
        tk.Scale(win, from_=0, to=10, resolution=0.5, orient='horizontal', variable=fade_out_var,
                 bg=PANEL, fg=ACCENT, troughcolor=BG, highlightthickness=0, length=220).pack(padx=18)

        def apply_and_close():
            self.fade_in_sec = fade_in_var.get()
            self.fade_out_sec = fade_out_var.get()
            self._save_config()
            win.destroy()

        tk.Button(win, text='Done', command=apply_and_close, bg=ACCENT, fg='#ffffff',
                  relief='flat', padx=18, pady=6, font=('Segoe UI', 9, 'bold')).pack(pady=18)
        win.protocol('WM_DELETE_WINDOW', apply_and_close)

    # ---------- pad lists ----------

    def _refresh_lists_array(self):
        self.lists = [{'id': lid, 'name': entry['name']} for lid, entry in self.all_lists_data.items()]

    def _refresh_list_combo(self):
        names = [l['name'] for l in self.lists] + [NEW_LIST_SENTINEL]
        self.list_combo['values'] = names
        idx = next((i for i, l in enumerate(self.lists) if l['id'] == self.current_list_id), 0)
        if names:
            self.list_combo.current(idx)

    def _on_list_selected(self, _event):
        idx = self.list_combo.current()
        # A readonly Combobox keeps keyboard focus after a selection, which
        # would otherwise silently swallow every keyboard shortcut (space,
        # arrows, Enter, 1-0/Q-P) from here on — hand focus back to the
        # window itself so they keep working right away.
        self.root.focus_set()
        if idx == len(self.lists):
            self._refresh_list_combo()  # revert visible selection until a name is confirmed
            name = simpledialog.askstring('New list', 'Name for the new list:',
                                           initialvalue=f'List {len(self.lists) + 1}', parent=self.root)
            if name and name.strip():
                self._create_list(name.strip())
            return
        if 0 <= idx < len(self.lists):
            self._switch_list(self.lists[idx]['id'])

    def _create_list(self, name):
        list_id = f'list_{uuid.uuid4().hex[:10]}'
        self.all_lists_data[list_id] = {'name': name, 'pads': {}, 'pad_count': BASE_PAD_COUNT}
        self._refresh_lists_array()
        self._switch_list(list_id)

    def _switch_list(self, list_id):
        """Switching only changes which list is displayed/edited — playback
        (self.now_playing) is untouched, so audio keeps playing right through."""
        if list_id == self.current_list_id:
            return
        self.current_list_id = list_id
        self._load_pads_for_current_list()
        self._update_transport()
        self._refresh_list_combo()
        self._save_config()

    def rename_list(self):
        entry = self.all_lists_data.get(self.current_list_id)
        if not entry:
            return
        new_name = simpledialog.askstring('Rename list', 'List name:', initialvalue=entry['name'],
                                           parent=self.root)
        if new_name and new_name.strip():
            entry['name'] = new_name.strip()
            self._refresh_lists_array()
            self._save_config()
            self._refresh_list_combo()

    # ---------- transport controls ----------

    def on_transport_play_pause(self):
        record = self.now_playing
        if not record:
            return
        player = record['player']
        if player.is_paused():
            player.resume()
        elif player.is_playing():
            player.pause()
        if record['list_id'] == self.current_list_id:
            self._render_pad(record['index'])
        self._update_transport()

    def on_transport_stop(self):
        self.queue = None
        self._fade_out_and_stop_now_playing()

    def on_transport_next(self):
        """Manual skip: advance the queue if one is running, else jump to the
        next assigned pad after the current one."""
        record = self.now_playing
        if not record:
            return
        if self.queue and self.queue['list_id'] == record['list_id']:
            self._fade_out_player_record(record)
            if self.queue['order'][self.queue['pos']] == record['index']:
                self._advance_queue()
            else:
                self._play_queue_item(self.queue['list_id'], self.queue['order'][self.queue['pos']])
            return
        nxt = self._next_assigned_pad(record['list_id'], record['index'])
        if nxt:
            idx, name, path = nxt
            self._start_track(record['list_id'], idx, name, path, crossfade=True)

    def _reclaim_focus(self, event):
        widget = event.widget
        if isinstance(widget, (tk.Entry, tk.Text)):
            return  # genuine text entry — leave focus alone
        self.root.after(10, lambda: self._finish_reclaim_focus(widget))

    def _finish_reclaim_focus(self, _clicked_widget):
        current = self.root.focus_get()
        if current is None:
            return
        try:
            if current.winfo_toplevel() is not self.root:
                return  # a dialog (Settings / Play List / rename prompt) owns focus — leave it
        except tk.TclError:
            return
        if isinstance(current, (tk.Entry, tk.Text)):
            return
        self.root.focus_set()

    def _on_volume_change(self, _value):
        self.vol_pct_lbl.config(text=f'{self.master_volume.get()}%')
        record = self.now_playing
        if record is None or id(record['player']) in self.fade_jobs:
            return
        record['player'].set_volume(self.master_volume.get() * 10)

    def _on_key(self, event):
        focus = self.root.focus_get()
        # Ignore global shortcuts while a secondary window (Settings / Play List)
        # or a text/slider/dropdown control has focus.
        if focus is not None:
            if focus.winfo_toplevel() is not self.root:
                return
            if isinstance(focus, (tk.Entry, tk.Scale, ttk.Combobox)):
                return

        if event.keysym == 'space':
            idx = self.selected_index
            if (self.now_playing and self.now_playing['list_id'] == self.current_list_id
                    and self.now_playing['index'] == idx):
                self.on_transport_play_pause()  # pause/resume the playing (selected) pad
            else:
                pad = self.pads[idx]
                if pad.name and not pad.missing:
                    self.activate_pad(idx)       # play the selected pad's song
            return
        if event.keysym == 'Right':
            self._move_selection(1)
            return
        if event.keysym == 'Left':
            self._move_selection(-1)
            return
        if event.keysym == 'Down':
            self._move_selection(COLS)
            return
        if event.keysym == 'Up':
            self._move_selection(-COLS)
            return
        if event.keysym == 'Return':
            self.on_pad_click(self.selected_index)
            return
        key = event.keysym.upper()
        if key in KEYS:
            self.on_pad_click(KEYS.index(key))

    def _poll_playing(self):
        pythoncom.PumpWaitingMessages()

        for i, pad in enumerate(self.pads):
            if pad.name and not pad.missing:
                self._render_pad(i)

        record = self.now_playing
        if (record is not None and id(record['player']) not in self.fade_jobs
                and time.time() - record.get('started_at', 0) > 1.0):
            player = record['player']
            if player.opened and not player.is_playing() and not player.is_paused():
                # Playback finished naturally (grace period avoids mistaking
                # the brief 'Transitioning' state right after play() for this).
                queued = (self.queue is not None and self.queue['list_id'] == record['list_id'])
                if queued:
                    # An active queue drives advancement, overriding repeat/play-all.
                    # If the ended song is the tracked queue slot, advance; if the
                    # queue was applied while an off-queue song played, begin it now.
                    player.stop()
                    player.close()
                    self.now_playing = None
                    if record['list_id'] == self.current_list_id:
                        self._render_pad(record['index'])
                    if self.queue['order'][self.queue['pos']] == record['index']:
                        self._advance_queue()
                    else:
                        self._play_queue_item(self.queue['list_id'], self.queue['order'][self.queue['pos']])
                elif self.repeat_one:
                    # Replay the same track on the same player (repeat wins).
                    player.set_volume(self.master_volume.get() * 10)
                    player.play()
                    record['started_at'] = time.time()
                    pump_com(0.1)
                else:
                    next_target = self._next_assigned_pad(record['list_id'], record['index']) \
                        if self.play_all else None
                    player.stop()
                    player.close()
                    self.now_playing = None
                    if record['list_id'] == self.current_list_id:
                        self._render_pad(record['index'])
                    if next_target is not None:
                        nx_index, nx_name, nx_path = next_target
                        self._start_track(record['list_id'], nx_index, nx_name, nx_path, crossfade=False)

        self._update_transport()
        self.root.after(250, self._poll_playing)

    # ---------- persistence ----------

    def _save_config(self):
        if self.current_list_id is not None and self.current_list_id in self.all_lists_data:
            entry = self.all_lists_data[self.current_list_id]
            # Start from whatever's already stored (this preserves rows beyond
            # current_pad_count that are hidden, not deleted, by remove_pad_row)
            # and only sync the currently visible/editable range from self.pads.
            pads_data = dict(entry.get('pads', {}))
            for index in range(self.current_pad_count):
                pad = self.pads[index]
                key = str(index)
                if pad.name and pad.path:
                    pads_data[key] = {'name': pad.name, 'path': pad.path}
                else:
                    pads_data.pop(key, None)
            entry['pads'] = pads_data

        data = {
            '__settings__': {'fade_in': self.fade_in_sec, 'fade_out': self.fade_out_sec,
                             'play_all': self.play_all, 'repeat_one': self.repeat_one},
            '__current_list__': self.current_list_id,
            'lists': self.all_lists_data,
            'saved_queues': self.saved_queues,
        }
        try:
            with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
        except OSError:
            pass

    def _init_default_list(self):
        self.all_lists_data = {DEFAULT_LIST_ID: {'name': DEFAULT_LIST_NAME, 'pads': {}, 'pad_count': BASE_PAD_COUNT}}
        self.current_list_id = DEFAULT_LIST_ID

    def _load_pads_for_current_list(self):
        entry = self.all_lists_data.get(
            self.current_list_id, {'name': DEFAULT_LIST_NAME, 'pads': {}, 'pad_count': BASE_PAD_COUNT})
        self.current_pad_count = entry.get('pad_count', BASE_PAD_COUNT)
        pads_data = entry.get('pads', {})
        for index in range(MAX_PAD_COUNT):
            pad = self.pads[index]
            info = pads_data.get(str(index)) if index < self.current_pad_count else None
            if info:
                pad.name = info.get('name')
                pad.path = info.get('path')
                pad.missing = not (pad.path and os.path.exists(pad.path))
            else:
                pad.name = None
                pad.path = None
                pad.missing = False
        self._apply_pad_count()
        for i in range(self.current_pad_count):
            self._render_pad(i)
        # Keep the keyboard cursor on the playing pad if it's in this list, else clamp.
        if (self.now_playing and self.now_playing['list_id'] == self.current_list_id
                and self.now_playing['index'] < self.current_pad_count):
            self.selected_index = self.now_playing['index']
        elif self.selected_index >= self.current_pad_count:
            self.selected_index = self.current_pad_count - 1
        self._render_selection()

    def _load_config(self):
        if not os.path.exists(CONFIG_PATH):
            self._init_default_list()
            self._load_pads_for_current_list()
            return
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except (OSError, json.JSONDecodeError):
            self._init_default_list()
            self._load_pads_for_current_list()
            return

        settings = data.get('__settings__')
        if settings:
            self.fade_in_sec = float(settings.get('fade_in', self.fade_in_sec))
            self.fade_out_sec = float(settings.get('fade_out', self.fade_out_sec))
            self.play_all = bool(settings.get('play_all', self.play_all))
            self.repeat_one = bool(settings.get('repeat_one', self.repeat_one))

        sq = data.get('saved_queues')
        if isinstance(sq, dict):
            self.saved_queues = sq

        if data.get('lists'):
            self.all_lists_data = data['lists']
        else:
            # Legacy single-list format: top-level numeric keys were pad entries.
            reserved = ('__settings__', '__current_list__', 'lists', 'saved_queues')
            legacy_pads = {k: v for k, v in data.items() if k not in reserved}
            self.all_lists_data = {DEFAULT_LIST_ID: {'name': DEFAULT_LIST_NAME, 'pads': legacy_pads,
                                                       'pad_count': BASE_PAD_COUNT}}

        if not self.all_lists_data:
            self._init_default_list()
        else:
            saved_current = data.get('__current_list__')
            self.current_list_id = saved_current if saved_current in self.all_lists_data \
                else next(iter(self.all_lists_data))

        self._load_pads_for_current_list()

    def _on_close(self):
        self.player_a.close()
        self.player_b.close()
        self.root.destroy()


def main():
    if sys.platform != 'win32':
        print('SongWeaver uses Windows Media Player COM automation and only runs on Windows.')
        sys.exit(1)
    pythoncom.CoInitialize()
    try:
        root = tk.Tk()
        PadBoardApp(root)
        root.mainloop()
    finally:
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    main()
