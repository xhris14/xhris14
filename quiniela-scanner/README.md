# Quiniela Scanner

Put a Liga MX quiniela slip under your **Logitech C920s**. The app waits until the slip is still and in focus, reads the
9 L / E / V marks, reads the handwritten name, and types the ticket into your **open Excel workbook** (next free ticket,
name + picks 1–9). Runs on your Windows PC.

```
camera → find slip → straighten → read 9 marks → read name → write row in Excel → "remove slip" → next
```

## 1. Install (Windows, once)

1. Install **Python 3.10 or newer** from python.org (tick *Add python.exe to PATH*).
2. Double-click **`setup.bat`** (creates a private environment and installs the packages).
3. *(Optional, for automatic name reading)* create a file called **`.env`** in this folder containing one line:
   `ANTHROPIC_API_KEY=sk-ant-...`  (see §5). Without it you simply type the name.
4. Open your jornada workbook in Excel, then double-click **`run.bat`**.

Try it with no camera first: `run.bat --demo path\to\photo-of-a-slip.jpg` loops your photo as a fake camera.

## 2. Camera setup (this matters more than anything else)

* Mount the C920s **looking straight down**, roughly **25–30 cm** above the table, so a slip fills about a third of the picture.
  Keep the whole slip in view, **including the NOMBRE line**.
* **Plain surface, even light.** The photo you sent had a hard shadow across the table; a desk lamp from the side
  or a dark mat under the slip avoids it. Glare on glossy paper is the other enemy.
* Close other apps that use the webcam (Zoom, Teams, OBS) and open the C920s privacy shutter.
* If autofocus hunts, untick **Autofocus** in the app and set `manual_focus` (0–255) in `config.json`; **Camera settings**
  opens the Logitech/DirectShow dialog for exposure and brightness.
* Camera on the wrong index? Change **Camera #** and press *Restart camera* (`run.bat cameras` lists what Windows sees).

## 3. How it finds your Excel sheet

Nothing is hard-coded to column letters. On every sheet named `ENTRY*` it looks for a header row containing
**`NOMBRE`** followed by **`1 2 3 4 5 6 7 8 9`**, with **`#`** (the ticket number) just left of `NOMBRE` – exactly like your
`ENTRY PG1` sheet:

```
 #  | NOMBRE      | 1 | 2 | … | 9 | T      ← header row (each block repeats this, side by side)
241 | RECA 30 RRR | V | V | … | L | 5
```

* Each block it finds appears in the **Block** drop-down with its ticket range and how many tickets are still free
  (`run.bat excel` prints the same thing – a good first check that it understood your file).
* It writes **name + picks 1–9 only**. `T` and everything else is left alone.
* Target = the **first free ticket** (empty name and empty picks) in the selected block. Type a **Ticket #** to override.
  It re-reads Excel just before every write, so rows you typed by hand in the meantime are never overwritten.
* **Group tag**: see §4a.
* After opening a different workbook (next jornada), press **Connect / reconnect Excel**. The app stays bound to the
  workbook it connected to, so switching windows in Excel can't redirect a write.
* It uses Excel's own object model, so the file can stay open, AutoSave keeps working, and nothing is overwritten
  by a second program. If Excel says it is busy (a cell is in edit mode) press Esc in Excel and retry.

## 4. Daily use

1. Place a slip → the yellow outline turns green → the card on the right fills in (straightened slip with the detected
   marks circled, name, jornada).
2. **Clean scans enter themselves** after 3 seconds. **Esc** cancels the countdown; touching anything takes over.
   Otherwise press **Enter** to write, **Esc** to discard.
3. Remove the slip; the app re-arms by itself. **Undo last** clears the previous ticket from Excel.

Auto-enter only fires when *all* of these hold; otherwise it waits for you:

* every row has one clear mark, or is simply empty (see §4b). Double marks, faint or scribbled marks are shown in red and always wait for you,
* Claude read the name confidently (if you type or fix a name yourself, press **Enter** to write it), and the slip's jornada matches the workbook's
  (`JORNADA 10 PC.xlsx` → 10; a slip printed "JORNADA 12" is flagged),
* the target ticket exists and is free.

### 4a. Group tags (`RRR`, `E`, …)

Some neighbourhoods/groups add a letter code so they can find themselves on the printed sheet. In the **Group tag**
row type the letters (one letter is fine), choose **Before name** or **After name**, and tick **Add tag (F2)**:

* While the tag is **on**, the green line says `TAG ON: "RRR" after name`, and the card shows exactly what will be
  written (`Will be written as: CLEMENTE O. RRR`).
* Next person has no group? Press **F2** (or untick the box). The tag is added to nothing until you switch it on again.
* The switch **always starts OFF** when you open the app, so no one is tagged by accident. The letters and the
  before/after choice are remembered.
* If you switch the tag or change it while a slip is on screen, that slip waits for **Enter** instead of auto-entering.
* If the slip already has the tag as its own word at that end (`TONY 5000 RRR`) it is not added twice. A tag of `E`
  is still added to `JOSE` or `ELENA`, because those are letters inside a name, not a tag.

### 4b. Empty boxes become `F1`

If a player leaves a match unmarked, that pick cell is filled with **`F1` in font size 9** (the rest of the row keeps its normal
font). On the card the row shows `empty = F1`, and the history list shows `F1` for it.

* A slip with 1–2 empty boxes still auto-enters. **3 or more empty boxes wait for you** (that usually means the camera missed
  something, e.g. a blank or badly lit slip). Change this with `auto_enter_max_blank`, or set `auto_enter_blanks` to `false` to
  always review slips that have an empty box.
* If you click **F1** on a row yourself, the same happens.
* **Undo last** removes the `F1` *and* puts the cell's previous font size back.
* Don't want the rule? Set `blank_marker` to `""` in `config.json` and empty boxes stay empty.

**`scans/`** keeps a picture of every
entered slip plus `log.csv`, so any dispute can be checked against what the camera saw.

## 5. Reading the handwritten name

Handwriting is where ordinary OCR fails, so the name is read with Claude vision (`claude_model` in `config.json`,
default `claude-opus-5-5`, low effort). Only **two small crops** – the printed title strip and the NOMBRE strip – are
sent, never the whole slip. Cost is a fraction of a cent per slip. No key / no internet / `read_names_with_claude: false`
→ the name box is simply left for you to type, and everything else works the same. Always glance at the name before it
goes in; Claude flags names it is unsure about and those never auto-enter.

## 6. Settings (`config.json`, created when you close the app)

| key | meaning |
|---|---|
| `camera_index`, `camera_backend` | which webcam; `dshow` (default, best for C920s) / `msmf` |
| `autofocus`, `manual_focus` | focus mode and the manual value (0–255) |
| `matches_per_slip` | rows on the slip (9) |
| `workbook` | blank = the workbook active in Excel at connect time, or part of a file name |
| `sheets` | sheet-name patterns to search (`["ENTRY*"]`) |
| `name_header` | header text above the names (`NOMBRE`) |
| `name_tag`, `name_tag_position`, `tag_separator` | remembered group-tag letters, `before` / `after`, and what goes between tag and name (default one space) |
| `uppercase_names` | write names in capitals |
| `auto_enter`, `auto_enter_delay_s` | hands-free entry and the cancel window |
| `blank_marker`, `blank_font_size` | text and font size for an empty box (`"F1"`, `9`; `""` = leave empty) |
| `auto_enter_blanks`, `auto_enter_max_blank` | whether empty boxes may auto-enter, and how many (default yes, up to 2) |
| `save_after_write` | also press Save in Excel after each ticket (AutoSave normally covers this) |
| `expected_jornada` | blank = take it from the workbook name |
| `min_ink`, `multi_ratio` | mark sensitivity – run `run.bat scan photo.jpg` to see the ink readings per cell |

## 7. Command line

```
run.bat                          the scanner window
run.bat --demo slip.jpg          same, with a photo as the "camera"
run.bat --demo slip.jpg --xlsx pool.xlsx   edit a closed .xlsx file instead of live Excel
run.bat scan slip.jpg --out r.jpg [--name]   read one photo, print picks (+ name with Claude), save the annotated image
run.bat excel                    show how the open workbook was understood
run.bat cameras                  list webcams
```

## 8. Troubleshooting

| symptom | fix |
|---|---|
| "cannot open camera" | shutter closed / another app using it / try Camera # 1 |
| "Slip seen but not recognised" | whole table incl. NOMBRE line must be in view and flat; reduce glare; check focus |
| a mark is missed or invented | `run.bat scan photo.jpg` prints ink per cell; adjust `min_ink` (lower = more sensitive) |
| "Excel is not running" | open the workbook first, then *Connect / reconnect Excel* |
| "no 'NOMBRE | 1 2 … 9' header found" | header text differs – set `name_header` / `sheets` in `config.json` |
| names always need typing | no `ANTHROPIC_API_KEY` in `.env`, or no internet |

## 9. What was tested, and what was not

Tested (automated, 46 tests): the slip reader on your real photo (all 9 marks, row for row) and on 13 degraded
versions of it (small, rotated ±25°, upside-down, blurred, noisy, dark, no blue pen, perspective); empty/unrelated pictures are
rejected; the capture logic (waits for stillness, reads once, needs the slip removed); Excel block discovery, free-ticket
choice, writing, undo, formula-safe names – against a workbook laid out like your screenshot; and the real window driven end to
end (manual entry, auto-enter, the safety rails above).

**Not tested on real hardware** – I had no Windows PC, webcam, Excel or API key in the build environment:
the live **C920s** capture, the **Excel COM connection** (verified only against a stand-in for Excel's object model), and the
**live Claude name read** (request shape verified against a mock). Do a first run on a *copy* of your workbook.

Developers: `pip install -r requirements-dev.txt`, then `pytest` (GUI tests need a display; on Linux use `xvfb-run pytest`).
