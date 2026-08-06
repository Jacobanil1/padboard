"""PadBoard - a 20-pad local soundboard for Windows.

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
ICON_PATH = os.path.join(_resource_dir(), 'padboard_icon.ico')

BG = '#ffffff'
PANEL = '#f2f7f2'
PANEL_BORDER = '#e0e8e1'
PAD_BG = '#43a047'
PAD_BORDER = '#2e7d32'
PAD_TEXT = '#ffffff'
PAD_TEXT_DIM = '#dcedc8'
ACCENT = '#16a34a'
ACCENT2 = '#4ade80'
PLAYING = '#ffd54f'
TEXT = '#16281c'
TEXT_DIM = '#5c7263'
DANGER = '#dc2626'
SLIDER_BLUE = '#2563eb'
SLIDER_BLUE_ACTIVE = '#3b82f6'

PAD_W = 170
PAD_H = 130
CORNER_W_FRAC = 0.32
CORNER_H_FRAC = 0.24

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

        self._build_ui()
        self._load_config()
        self._refresh_lists_array()
        self._refresh_list_combo()
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)
        self.root.bind_all('<Key>', self._on_key)
        self._update_transport()
        self._poll_playing()

    # ---------- UI ----------

    def _build_ui(self):
        self.root.title('PadBoard - Local Soundboard')
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

        title = tk.Label(header, text='● PadBoard', font=('Segoe UI', 16, 'bold'),
                          bg=BG, fg=TEXT)
        title.pack(side='left')

        settings_btn = tk.Button(header, text='⚙ Settings', command=self.open_settings,
                                  bg=PANEL, fg=TEXT, activebackground=PAD_BORDER, activeforeground=TEXT,
                                  relief='flat', padx=14, pady=7, font=('Segoe UI', 9))
        settings_btn.pack(side='right')

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

        self.now_icon_lbl = tk.Label(now_frame, text='♪', font=('Segoe UI', 13), bg=PAD_BG, fg='#ffffff',
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

        self.play_pause_btn = tk.Button(btns, text='▶', command=self.on_transport_play_pause,
                                         bg=ACCENT, fg='#ffffff', activebackground=ACCENT2,
                                         activeforeground='#ffffff', relief='flat', bd=0,
                                         width=3, height=1, font=('Segoe UI', 18, 'bold'),
                                         disabledforeground='#ffffff', state='disabled', cursor='hand2')
        self.play_pause_btn.pack(side='left', padx=(0, 10))

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

        self.seek_duration_lbl = tk.Label(seek_frame, text='0:00', font=('Consolas', 8), bg=PANEL, fg=TEXT_DIM,
                                           width=4)
        self.seek_duration_lbl.pack(side='left')

        self.stop_btn = tk.Button(btns, text='⏹', command=self.on_transport_stop,
                                   bg=BG, fg=TEXT, activebackground=PANEL_BORDER, activeforeground=TEXT,
                                   relief='flat', bd=1, width=3, height=1, font=('Segoe UI', 15),
                                   state='disabled', cursor='hand2')
        self.stop_btn.pack(side='left')

        # Volume (right)
        vol_frame = tk.Frame(inner, bg=PANEL)
        vol_frame.pack(side='right')

        tk.Label(vol_frame, text='🔊', font=('Segoe UI', 12), bg=PANEL, fg=TEXT_DIM).pack(side='left', padx=(0, 8))
        vol_scale = tk.Scale(vol_frame, from_=0, to=100, orient='horizontal', variable=self.master_volume,
                              bg=PANEL, fg=SLIDER_BLUE, troughcolor='#dfe8e0', highlightthickness=0,
                              length=150, sliderlength=18, showvalue=False, bd=0,
                              activebackground=SLIDER_BLUE_ACTIVE, command=self._on_volume_change)
        vol_scale.pack(side='left')
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
                                      font=('Segoe UI', 12, 'bold'), width=w - 22,
                                      justify='center', tags=('padclick',))

        clear_id = canvas.create_text(w - 14, h - 12, text='', fill=PAD_TEXT_DIM,
                                       font=('Segoe UI', 9), tags=('clearbtn',))

        canvas.tag_bind('padclick', '<Button-1>', lambda e, i=index: self._on_pad_canvas_click(e, i))
        canvas.tag_bind('clearbtn', '<Button-1>', lambda e, i=index: self.clear_pad(i))
        canvas.tag_bind(name_id, '<Double-Button-1>', lambda e, i=index: self.rename_pad(i))
        canvas.bind('<Enter>', lambda e, i=index: self._pad_hover(i, True))
        canvas.bind('<Leave>', lambda e, i=index: self._pad_hover(i, False))

        return {
            'canvas': canvas, 'bg': bg_id, 'gloss': gloss_id, 'key': key_id, 'num': num_id,
            'icon_bg': icon_bg_id, 'icon': icon_id, 'name': name_id, 'clear': clear_id,
        }

    def _pad_hover(self, index, entering):
        w = self.pad_widgets[index]
        w['canvas'].itemconfig(w['bg'], fill=(lighten(PAD_BG, 0.08) if entering else PAD_BG))

    def _on_pad_canvas_click(self, event, index):
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

        if pad.name:
            c.itemconfig(w['icon'], text=('⏸' if playing else '▶'), fill=PAD_TEXT)
            display_name = pad.name if not pad.missing else f'{pad.name} (missing)'
            c.itemconfig(w['name'], text=display_name, fill=(DANGER if pad.missing else PAD_TEXT))
            c.itemconfig(w['clear'], text='✕')
            border = PLAYING if is_current else PAD_BORDER
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
            self._update_seek()
            return
        paused = record['player'].is_paused()
        list_name = self.all_lists_data.get(record['list_id'], {}).get('name', '')
        prefix = f'{list_name}  ·  ' if list_name else ''
        self.now_icon_lbl.config(text=('▶' if paused else '⏸'))
        self.now_title_lbl.config(text=record['name'])
        self.now_sub_lbl.config(text=f"{prefix}Pad {record['index'] + 1}  ·  " + ('paused' if paused else 'playing'))
        self.play_pause_btn.config(text=('▶' if paused else '⏸'), state='normal')
        self.stop_btn.config(state='normal')
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
            messagebox.showerror('PadBoard', f"Couldn't load this file:\n{os.path.basename(path)}\n\n{err_msg}\n\n"
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

        if self._is_now_playing(index):
            self._fade_out_and_stop_now_playing()
            return

        previous = self.now_playing
        incoming = self.player_b if (previous and previous['player'] is self.player_a) else self.player_a

        ok, err_msg = incoming.open(pad.path)
        if not ok:
            messagebox.showerror('PadBoard', f"Couldn't play this file:\n{pad.name}\n\n{err_msg}")
            return

        incoming.set_volume(0)
        incoming.play()
        pump_com(0.15)
        self.now_playing = {'list_id': self.current_list_id, 'index': index, 'name': pad.name,
                             'player': incoming, 'started_at': time.time()}
        if previous is not None:
            self._fade_out_player_record(previous)
        self._fade_in_player(incoming)
        self._render_pad(index)
        self._update_transport()

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
        self._fade_out_and_stop_now_playing()

    def _on_volume_change(self, _value):
        self.vol_pct_lbl.config(text=f'{self.master_volume.get()}%')
        record = self.now_playing
        if record is None or id(record['player']) in self.fade_jobs:
            return
        record['player'].set_volume(self.master_volume.get() * 10)

    def _on_key(self, event):
        focus = self.root.focus_get()
        if isinstance(focus, (tk.Entry,)):
            return
        if event.keysym == 'space':
            self.on_transport_play_pause()
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
                player.stop()
                player.close()
                self.now_playing = None
                if record['list_id'] == self.current_list_id:
                    self._render_pad(record['index'])

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
            '__settings__': {'fade_in': self.fade_in_sec, 'fade_out': self.fade_out_sec},
            '__current_list__': self.current_list_id,
            'lists': self.all_lists_data,
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

        if data.get('lists'):
            self.all_lists_data = data['lists']
        else:
            # Legacy single-list format: top-level numeric keys were pad entries.
            legacy_pads = {k: v for k, v in data.items() if k not in ('__settings__', '__current_list__', 'lists')}
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
        print('PadBoard uses Windows Media Player COM automation and only runs on Windows.')
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
