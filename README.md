# Interview Analysis Pipeline

Paste a YouTube link to a recorded interview → get an Excel file with every question as a column and the condensed, translated answer underneath.

Everything runs through the [Groq API](https://console.groq.com) (free tier). No GPU, no local Whisper model, no paid subscriptions.

---

## What it produces

An Excel file with two sheets:

**Sheet 1 — Summary**  
Each interviewer question is a column header. The row below contains the condensed, translated answer. A second row shows the original-language answer for reference.

**Sheet 2 — Full detail**  
One row per question. Columns: question number, timestamp, language, original question, translated question, condensed answer (original), condensed answer (translated).

---

## Requirements

**Python 3.8 or higher**

Install the Python packages:
```bash
pip install groq openpyxl yt-dlp
```

Install ffmpeg (required for audio conversion):

| OS | Command |
|---|---|
| Windows | `winget install ffmpeg` |
| Mac | `brew install ffmpeg` |
| Linux | `sudo apt install ffmpeg` |

Get a free Groq API key at [console.groq.com](https://console.groq.com) → API Keys → Create key. No credit card needed.

---

## Setup

1. Clone or download this repository
2. Install the packages above
3. Open `interview_pipeline_v7.py` and fill in the `CONFIG` section at the top

---

## Configuration

Open the file and edit the `CONFIG` section near the top. These are the only lines you need to change:

```python
# Your Groq API key
GROQ_API_KEY = "paste_your_key_here"

# YouTube URL of the interview
YOUTUBE_URL = "https://youtube.com/live/..."

# Language for the Excel output
# 'ru' = Russian | 'uz' = Uzbek | 'en' = English
OUTPUT_LANGUAGE = "ru"

# Language spoken in the video
# 'ru' = Russian | 'uz' = Uzbek | None = auto-detect
AUDIO_LANGUAGE = "ru"

# Where to save output files (created automatically)
OUTPUT_DIR = "./output"

# Path to your ffmpeg folder (Windows only — set to None if ffmpeg is on PATH)
FFMPEG_DIR = r"C:\Users\YourName\ffmpeg\bin"
```

**`FFMPEG_DIR`** — on Windows, if `ffmpeg` is not on your system PATH, set this to the folder containing `ffmpeg.exe` and `ffprobe.exe`. On Mac and Linux, set it to `None`.

---

## Usage

```bash
python interview_pipeline_v7.py
```

The script prints progress as it runs:

```
============================================================
Interview Analysis Pipeline — v7
Output language : Russian
Audio language  : ru
Output folder   : C:\Users\...\output
============================================================

[1/5] Downloading audio...
Title    : Кредитные Карты_Глубинное интервью № 1
Duration : 46m 20s

[2/5] Splitting into chunks...
Splitting into 6 chunks of ~8 min...

[3/5] Transcribing with whisper-large-v3-turbo...
Chunk 1/6 (12.3 MB)... 42 segments | lang: ru | 1.4s
...

[4/5] Building transcript...
[5/5] Extracting Q&A with llama-3.3-70b-versatile...

All done. Files in: C:\Users\...\output
```

Output files are saved to `OUTPUT_DIR`:
- `interview_analysis.xlsx` — the Excel file
- `raw_transcript.txt` — the full timestamped transcript from Whisper

---

## Supported YouTube URL formats

All of these work:

```
https://youtube.com/live/xxxxxxxxxxx?feature=share
https://www.youtube.com/watch?v=xxxxxxxxxxx
https://youtu.be/xxxxxxxxxxx
https://www.youtube.com/shorts/xxxxxxxxxxx
```

---

## Troubleshooting

**"Download failed" / "Sign in to confirm you're not a bot"**  
YouTube is blocking the download. Install the [Get cookies.txt LOCALLY](https://chromewebstore.google.com/detail/cclelndahbckbenkjhflpdbgdldlbecc) Chrome extension, go to youtube.com while logged in, click Export, then add this to the yt-dlp call in `get_yt_audio()`:
```python
dl_cmd += ["--cookies", "path/to/youtube.com_cookies.txt"]
```

**"File too large"**  
Reduce the chunk size. In `split_audio()`, change `chunk_minutes=8` to `chunk_minutes=5`.

**Transcript looks garbled ("Q. Q. Q." repeated)**  
`AUDIO_LANGUAGE` is set wrong. Change it to match the language actually spoken in the video (`'ru'` for Russian, `'uz'` for Uzbek).

**Only the last few minutes were extracted**  
Rate limits cut off some LLM chunks. Wait 2-3 minutes and run the script again.

**Nothing extracted at all**  
Open `output/raw_transcript.txt` and check whether the text makes sense. If it is mostly noise or the wrong language, `AUDIO_LANGUAGE` needs to be corrected or the audio quality is too low.

**Mixed languages in the Excel (some columns Russian, some English)**  
The script includes an automatic repair pass that detects and re-translates any fields left in the wrong language. If it still happens, run the script again.

**ffmpeg not found (Windows)**  
Set `FFMPEG_DIR` in the config to the full path of your ffmpeg `bin` folder:
```python
FFMPEG_DIR = r"C:\Users\YourName\ffmpeg\ffmpeg-master-latest-win64-gpl\bin"
```

---

## Security note

**Do not commit your API key to Git.** Keep the `GROQ_API_KEY` out of version control. The recommended approach is a `.env` file:

```bash
pip install python-dotenv
```

Create a `.env` file in the same folder (and add it to `.gitignore`):
```
GROQ_API_KEY=your_key_here
```

Then at the top of the script, replace the hardcoded key with:
```python
from dotenv import load_dotenv
load_dotenv()
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
```

---

## How it works

```
YouTube URL
  → download audio as 16kHz mono FLAC        (yt-dlp + ffmpeg)
  → split into 8-minute chunks               (ffmpeg)
  → transcribe each chunk                    (Groq Whisper Large v3 Turbo)
  → format into timestamped transcript
  → extract Q&A + translate                  (Groq Llama 3.3 70B)
  → export to Excel                          (openpyxl)
```

**Why 16kHz mono FLAC?**  
16kHz is Whisper's native sample rate — higher wastes space. Mono halves file size. At these settings 8 minutes of audio ≈ 10-12 MB, keeping chunks under Groq's 25 MB API limit.

**Why split into chunks?**  
A 46-minute interview produces an ~85 MB file. Groq rejects anything over 25 MB. Chunks are 8 minutes with an 8-second overlap so no words are lost at boundaries.

**Models used (both free)**  
- Transcription: `whisper-large-v3-turbo` — Groq runs this at 216x real-time. A 1-hour interview transcribes in ~15 seconds.  
- Q&A extraction: `llama-3.3-70b-versatile` — 70B parameter model, best available on Groq's free tier.

---

## Interview context

The script is configured for Russian-language qualitative research interviews about credit card usage in Uzbekistan. The interviewer is named Zina.

To adapt it for a different topic or interviewer, update the system prompt inside `build_qa_prompt()`.
