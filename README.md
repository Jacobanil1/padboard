# PadBoard

A 20-100 pad local soundboard for Windows. Assign audio files from your library to pads
and trigger them by click or keyboard, with crossfade, per-list pad banks, and a seek bar.

Two versions are included:

- **`padboard.py`** — native Windows desktop app (Tkinter + Windows Media Player COM
  automation for playback). This is what the packaged `PadBoard.exe` release runs.
- **`PadBoard.html`** — a self-contained browser version. No install needed, just open
  it in Edge or Chrome. Songs are stored in the browser's local database.

## Install (Windows)

Download the latest `PadBoard.exe` from [Releases](../../releases) and run it directly —
no installer, no Python required.

Or via [winget](https://learn.microsoft.com/windows/package-manager/winget/):

```
winget install Jacobanil1.PadBoard
```

## Features

- 20 pads per list by default, expandable up to 100 (5 more per click)
- Multiple named pad lists/banks, switchable without interrupting playback
- Crossfade between pads with adjustable fade-in/fade-out timing
- Seek bar, play/pause/stop transport controls, volume control
- Keyboard shortcuts (1-0, Q-P) for the first 20 pads
- Supports MP3, WAV, WMA, and M4A (via Windows Media Player's codecs)

## Running from source

Requires Python 3 on Windows, with `pywin32` installed:

```
pip install pywin32
python padboard.py
```

## License

MIT — see [LICENSE](LICENSE).
