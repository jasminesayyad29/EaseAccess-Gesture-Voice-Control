# Omega Voice Assistant Documentation

## Scope
This document explains how [src/Omega.py](src/Omega.py) works, step by step, including:
- Libraries used and why they are needed
- End-to-end runtime flow (startup, listening, intent routing)
- OCR-based click/select/copy/paste workflows
- Presentation and browser/tab controls
- Emoji-based intent and action mapping

## Main File
- [src/Omega.py](src/Omega.py)

## Related Files
- [src/app.py](src/app.py)
- [requirements.txt](requirements.txt)

## Libraries Used

### Voice Input/Output
- `pyttsx3`: text-to-speech (assistant voice replies)
- `speech_recognition`: microphone capture and Google STT recognition
- `datetime`, `time`, `date`: time/date queries, delays, scheduling behavior

### Input Automation and UI Control
- `pynput.keyboard`: keyboard press/hotkey simulation
- `pyautogui`: mouse movement/click/scroll, key presses, hotkeys
- `pygetwindow`: window discovery/focus/minimize/maximize
- `webbrowser`: open web and maps search URLs

### OCR and Vision
- `pytesseract`: OCR text detection from screen captures
- `PIL.ImageGrab`, `PIL.Image`, `ImageEnhance`, `ImageOps`: screenshots and OCR preprocessing
- `cv2`, `numpy`: template matching and image transforms for robust visual matching

### App/OS Integration (Windows)
- `os`, `glob`, `subprocess`, `psutil`, `winreg`: app discovery and process/registry lookup
- `win32gui`, `win32con`, `win32api`, `pythoncom`: low-level Windows window/taskbar interactions
- `pychrome`: browser DevTools tab listing and tab switching helpers

### NLP / Intent Recognition
- `re`, `json`, `difflib.SequenceMatcher`, `difflib.get_close_matches`: text normalization and fuzzy matching
- `scikit-learn` (`TfidfVectorizer`, `MultinomialNB`, `make_pipeline`): ML intent classifier
- `joblib`: model utilities (prepared for persistence workflows)

## Core Runtime State
Key global states maintained during runtime:
- `is_awake`: sleep/wake state of assistant
- `last_context`: remembers prior context (for example, presentation)
- `last_found_boxes`: OCR match boxes for disambiguation
- `last_tab_boxes`, `last_tab_target`: tab selection disambiguation
- `paste_word_mode`, `paste_word_boxes`, `paste_word_target`: paste-before/after flow
- `paste_position_mode`, `paste_position_boxes`: generic paste-position workflow
- `range_mode`, `range_phase`, related range variables: select/copy from phrase A to phrase B

## Step-by-Step Working

### 1) Startup
- Initializes recognizer and voice engine
- Tunes STT thresholds for noisy rooms
- Trains the intent model from built-in command examples
- Starts chat UI thread through `app.ChatBot.start`
- Greets user using `wish()`

### 2) Input Loop
Main loop continuously checks:
- GUI text input from chatbot UI
- Or microphone input via `record_audio()`

Input is normalized and wake-word gated:
- Wake variants: omega, oh mega, o mega, ome ga, amiga, omegaa

### 3) Speech-to-Text Robust Selection
`record_audio()` uses `_best_transcript_from_google()`:
- requests multiple STT alternatives
- scores candidates by command similarity + STT confidence
- returns best normalized transcript

### 4) Intent Detection
`respond()` sends cleaned text to `IntentRecognizer.predict_intent()`:
- exact phrase map first
- ML prediction next (TF-IDF + Naive Bayes)
- fallback keyword matcher if confidence is low

### 5) Entity Extraction
`extract_entity()` pulls command targets by intent, for example:
- search query text
- location/place name
- app name
- click target text
- tab name
- range commands (`from ... to ...`)

### 6) Action Routing
`respond()` routes to dedicated handlers:
- greeting/time/date/name
- search/maps
- files/folders navigation and opening
- app launching and taskbar search
- click/double-click actions
- copy/paste/range selection workflows
- tab switching and previous tab return
- presentation controls

### 7) OCR-Driven Operations
OCR workflows are used for:
- finding text on screen (`find_text_boxes`)
- selecting among multiple matches with numbered overlays
- selecting ranges (`select from ... to ...`)
- copying ranges (`copy from ... to ...`)
- smart pasting before/after words
- choosing exact paste positions

### 8) Browser and Presentation Context Handling
- Search can remember if user came from presentation context
- Browser windows can be minimized before returning to slides
- PowerPoint focus logic attempts slideshow window first, then editor, then fallback

### 9) Shutdown
- Handles `SystemExit`, `KeyboardInterrupt`, and unexpected errors safely
- Speaks status updates where possible

## Intent and Action Map (Emoji)

### Emoji Legend
- 🎤 voice input
- 🧠 intent detection
- 🗣️ assistant speech output
- 🌐 web action
- 📍 maps/location
- 📂 file/folder action
- 🪟 window/app action
- 🖱️ mouse click action
- ⌨️ keyboard shortcut action
- 🧾 OCR text/box action
- 📋 clipboard action
- 📊 presentation action
- 🔄 workflow state/disambiguation

### General Assistant Intents

| Intent | Typical Command | Action |
|---|---|---|
| 👋 `greeting` | hello / hi | 🗣️ Greets user via `wish()` |
| 🪪 `name_query` | what is your name | 🗣️ Replies assistant identity |
| 🕒 `time_query` | what time is it | 🗣️ Speaks current time |
| 📅 `date_query` | what is the date | 🗣️ Speaks current date |
| 😴 `wake_up` | wake up omega | 🔄 Sets awake state and greets |
| 👋 `goodbye` | goodbye / exit | 🔄 Sets sleep state and replies |

### Web and Location Intents

| Intent | Typical Command | Action |
|---|---|---|
| 🔎 `search` | search for climate change | 🌐 Opens Google search |
| 📍 `location` | locate Mumbai | 🌐 Opens Google Maps place URL |

### File and App Intents

| Intent | Typical Command | Action |
|---|---|---|
| 📂 `files` | open downloads | 📂 Opens folders/files by friendly names, full paths, or recursive search |
| ↩️ `file_navigate` | go back | 📂 Moves to parent directory |
| 🚀 `open_app` | open chrome | 🪟 Tries direct launch, taskbar OCR hover, then template matching |

### Click / OCR / Choice Intents

| Intent | Typical Command | Action |
|---|---|---|
| 🖱️ `click_action` | click submit | 🧾 OCR find target then 🖱️ click/double click |
| 🔢 `choice_select` | choose 2 | 🔄 Resolves multi-match selection (tab/click/range/paste flows) |
| 📑 `tab_control` | switch to tab docs | 🪟 Switches Chrome tab by OCR tab-text matching |

### Clipboard and Text-Manipulation Intents

| Intent | Typical Command | Action |
|---|---|---|
| 📋 `copy` | copy | ⌨️ Ctrl+C |
| 🧾➡️📋 `copy from ... to ...` | copy from intro to summary | 🧾 OCR range select + ⌨️ Ctrl+C |
| 📋➡️ `paste` | paste | ⌨️ Ctrl+V |
| 📋➡️➡️ `paste_after` | paste after budget | 🧾 OCR locate phrase and paste after |
| ⬅️📋 `paste_before` | paste before total | 🧾 OCR locate phrase and paste before |
| ⬆️📋 `paste_line_start` | paste at start of line heading | 🧾 Move caret to line start and paste |
| ⬇️📋 `paste_line_end` | paste at end of paragraph note | 🧾 Move caret to line end and paste |
| 🧮📋 `paste_cell` | paste in cell B12 | ⌨️ Excel Go To (Ctrl+G) then paste |
| 🧾 `range_select` | select from alpha to beta | 🧾 OCR range selection without copy |

### Presentation Intents (PowerPoint)

| Command Pattern | Action |
|---|---|
| ⏭️ next / forward / advance | 📊 Next slide (Right Arrow) |
| ⏮️ previous / back | 📊 Previous slide (Left Arrow) |
| ▶️ start / begin / slideshow | 🪟 Focus PowerPoint then start slideshow (F5) |
| ⏹️ end / exit / stop | 🪟 Return focus and exit slideshow (Esc) |
| ⏸️ pause | 📊 Black screen toggle (`b`) |
| ▶️ resume / continue | 📊 Black screen toggle (`b`) |
| 🔍➕ zoom in | ⌨️ Ctrl + `+` |
| 🔍➖ zoom out | ⌨️ Ctrl + `-` |
| 🖥️ full screen | ⌨️ F11 |
| ↩️ exit full screen | ⌨️ Esc |
| ⏮️ first slide | ⌨️ Home |
| ⏭️ last slide | ⌨️ End |
| 🔢 go to slide N | ⌨️ Type number + Enter |
| 🖍️ highlight/find phrase | ⌨️ Ctrl+F + phrase |
| ⬇️ scroll down | 🖱️ Scroll down |
| ⬆️ scroll up | 🖱️ Scroll up |

## Advanced Matching Workflows

### App Launch Strategy (3 Layers)
1. Direct app launch using aliases/executables
2. Taskbar icon centers + hover OCR tooltip detection
3. Template matching against screen image

### OCR Matching Strategy
- Word-level exact/fuzzy matching
- Token normalization and OCR confusion fixes
- Cluster/line-level phrase grouping
- Chrome-specific second-pass enhanced OCR

### Disambiguation Pattern
When multiple matches are found:
- Overlay numbered options on screen
- Assistant prompts: choose 1 / choose 2 / ...
- `handle_choice_selection()` or specialized range/paste choice handlers finalize action

## Important Functions (Quick Reference)
- `respond(voice_data)`: central command router
- `IntentRecognizer`: ML + fallback intent system
- `extract_entity(text, intent)`: pulls command target
- `find_text_boxes(target)`: OCR matching engine
- `perform_click_action(entity, ...)`: click/double-click logic
- `start_range_selection(...)`: range selection/copy workflow
- `handle_paste_after(...)`: smart paste before/after handling
- `open_system_app(...)` and `open_taskbar_app(...)`: app launching
- `handle_presentation_control(command)`: slideshow controls

## Run Instructions
From project root:

```bash
python src/Omega.py
```

## Common Troubleshooting
- 🎤 Microphone not capturing:
  - check Windows microphone permissions
  - verify input device and default microphone
- 🗣️ No voice output:
  - verify SAPI voice availability (`pyttsx3`)
- 🧾 OCR misses text:
  - improve on-screen contrast/zoom
  - ensure Tesseract path is correct (`pytesseract.pytesseract.tesseract_cmd`)
- 🪟 App launch not reliable:
  - app may not be indexed in Start Menu/Registry
  - try explicit app name variants (for example `open app chrome`)
- 📊 Presentation focus issues:
  - ensure PowerPoint window is not blocked/minimized behind other windows

## Summary
[src/Omega.py](src/Omega.py) is a full desktop voice automation controller that combines:
- 🎤 robust speech recognition
- 🧠 ML + rule-based intent detection
- 🧾 OCR-driven visual targeting
- 🖱️ input automation for apps, browser tabs, and presentations
- 🔄 multi-step disambiguation workflows for reliable command execution
