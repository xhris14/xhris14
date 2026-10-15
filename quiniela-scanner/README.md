# Quiniela Scanner

Put a stack of Liga MX quiniela slips under your **Logitech C920s** and take them off one at a time. The app waits until
the top slip is still and in focus, reads the 9 L / E / V marks, **types the picks into your open Excel workbook at once**,
and fills in the handwritten name a few seconds later. Lift the slip, and the next one (now on top, in the same place) is
read and entered by itself. Runs on your Windows PC.

```
camera → top slip still? → read 9 marks → picks into Excel (next free ticket) → you lift the slip → next slip
                                              └→ name read in the background → name appears in the same row
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
* If autofocus hunts, untick **Autofocus** and use the **Focus** slider next to it (0–255, steps of 5; it changes the camera
  live and is remembered as `manual_focus`). **Camera settings** opens the Logitech/DirectShow dialog for exposure and brightness.
* **Stacking**: square the pile up so every new top slip lies in the same place, and keep the whole slip in view. Lift each slip
  completely off the pile (a slip that is only slid aside is treated as unchanged); a hand reaching in is what tells the app
  "the next one is coming".
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
* Target = the **first free ticket** (empty name and empty picks) in the selected block. Type a **Ticket #** to override
  (for that one slip).
* The sheets are read in full **once**, when you connect. After that every slip costs one small read of its target row
  (is it still that ticket, and still empty?) and one write, so Excel is never busy for long and rows you typed by hand in
  the meantime are never overwritten (the app moves on to the next free row). If rows were inserted or deleted above a block
  it notices by the ticket number and re-reads. A row you *clear* by hand is only seen as free again after **Connect /
  reconnect Excel**.
* **Group tag**: see §4a.
* After opening a different workbook (next jornada), press **Connect / reconnect Excel** (the only button that re-reads
  everything). The app stays bound to the workbook it connected to, so switching windows in Excel can't redirect a write.
* It uses Excel's own object model, so the file can stay open, AutoSave keeps working, and nothing is overwritten
  by a second program. If Excel says it is busy (a cell is in edit mode) press Esc in Excel and retry.

## 4. Daily use (the stack workflow)

1. Put the stack under the camera and tick **Auto-enter clean scans** (on by default). Settle the top slip: the yellow outline
   turns green and, within about half a second, the picks are in Excel (ticket number shown on the card:
   **ENTERED - ticket 251**, plus a short beep on Windows).
2. **Lift the slip.** The next slip is read and entered the same way. Nothing waits for the name and nothing needs a click.
3. A few seconds after each slip (as long as Claude takes to answer) the name appears in its row. Several names can be in
   flight at once and come back in any order; each lands in its own row. The strip under the preview shows how many slips you entered, **seconds per slip** (average of
   the last 10; pauses over a minute are ignored), and the last slip's times: read, Excel, name.

A slip goes in **at once** only when *all* of these hold; otherwise it opens the classic card and waits for you
(and you hear the "attention" beep):

* Auto-enter is ticked, `auto_enter_delay_s` is 0, Excel is connected and a name reader is available (`ANTHROPIC_API_KEY`),
* every row has one clear mark, or is simply empty (see §4b). Double marks, faint or scribbled marks are shown in red.

**The classic card** (also used when Auto-enter is off or there is no API key): the slip, the detected marks and the name box
are shown; fix whatever is wrong and press **Enter** to write, **Esc** to discard. **The camera keeps reading while a card is open**:
every slip you lift and put down meanwhile waits in a queue and is entered, in order, as soon as you have finished the card. That
queue is loud, because those tickets are not in Excel yet: an attention beep when the first one queues up, a red bar ("2 slips are
WAITING behind the card ...") and "slips waiting behind the card: N" in the strip. Slips are never dropped for a card; closing the
window while the card or the queue still holds slips asks first. With `auto_enter_delay_s` above 0 you get the old hands-on
mode instead: wait for the name, count down that many seconds (**Esc** cancels), then write.

**Keys:** **Enter** writes the card and **Esc** discards it, but not while you are typing in the Ticket box, the Group-tag box, the
review name box or a drop-down: there they only finish that entry (Enter / Esc leave it; the next Enter writes the card). **F2**
switches the tag from anywhere. If a card still shows a row the camera could not read (two marks, faint, scribbled) on "empty" and
you press Enter, the app asks before writing `F1` for it.

**When the name is not certain** (history rows turn **red**):

| what happened | what is written in the name cell |
|---|---|
| Claude read it confidently | the name (with the group tag, §4a) |
| Claude was unsure | the best guess followed by ` ?` (`unsure_marker`) |
| read failed, timed out (45 s), declined, no internet | just `?` |
| somebody typed in that cell in Excel meanwhile | nothing - the row says **NAME NOT WRITTEN - type it** |

Click a red row (or press **Next to review**): the name picture appears under the preview with the guess in an entry
(marker removed). Correct it and press **Save name** (or Enter in that entry) - it is written, and the row stops being red.
The app only overwrites a name cell that is still empty or still what it wrote itself. **Undo this entry** clears that ticket.

**Wrong jornada** (the slip says JORNADA 12, the workbook is JORNADA 10): the row goes in, turns red as **WRONG JORNADA**,
you hear the attention beep and **Auto-enter is switched off** (circuit breaker) so the rest of the stack goes to the card
until you look, undo what is wrong and tick Auto-enter again. (The switch-off is not saved: next start it is on again.)
The jornada is the number in the workbook name (`JORNADA 10 PC.xlsx` → 10) unless `expected_jornada` says otherwise.

**Undo last** clears the previous ticket from Excel - also while its name is still being read (the late name is ignored).
A slip that is lifted and put back is **not** entered twice (the app recognises the pen marks and handwriting); each repeat it drops is
counted (**Skipped repeats: N** in the strip, `DUPLICATE_SKIPPED` in `scanner.log`) and "Same slip still there - lift it off." stays on
screen while that slip lies there. **Capture now** reads whatever is there, repeats included, so it ignores a second click within
1.5 s (a double click must not enter the slip twice) and does nothing while a card is open. Every entered slip is kept in `scans/`
with `log.csv` (see §6).

**Closing the window** waits up to 10 s for names that are still being read, then writes `?` for any that never came back, so every
slip has its picture and its `log.csv` row. The process exits at once, even if a name request is still hanging.

**Small screens:** the window is sized to the screen (1366x768, 1280x720, Windows display scaling of 125-150 %): the layout gets denser
and, if it is still too tall, the middle part scrolls while the strip, the history list, **Undo last** and **Next to review** stay in
view. **Jornada check: 12** (green) next to the group tag says the wrong-jornada guard is on; **Jornada check OFF** (amber) means the
workbook name has no `jornada <number>` (separators, `No.` and `#` are fine) - set `expected_jornada` in `config.json` then.

### 4a. Group tags (`RRR`, `E`, …)

Some neighbourhoods/groups add a letter code so they can find themselves on the printed sheet. In the **Group tag**
row type the letters (one letter is fine), choose **Before name** or **After name**, and tick **Add tag (F2)**:

* While the tag is **on**, the green line says `TAG ON: "RRR" after name`, and the card shows exactly what will be
  written (`Will be written as: CLEMENTE O. RRR`).
* Next person has no group? Press **F2** (or untick the box). The tag is added to nothing until you switch it on again.
* The switch **always starts OFF** when you open the app, so no one is tagged by accident. The letters and the
  before/after choice are remembered.
* The tag settings are **noted when the slip is scanned**, not when its name arrives: press F2 for the next person right
  away and the slip before still gets the setting it was scanned with. (On the classic card, changing the tag makes that
  slip wait for **Enter**.)
* If the slip already has the tag as its own word at that end (`TONY 5000 RRR`) it is not added twice. A tag of `E`
  is still added to `JOSE` or `ELENA`, because those are letters inside a name, not a tag.

### 4b. Empty boxes become `F1`

If a player leaves a match unmarked, that pick cell is filled with **`F1` in font size 9** (the rest of the row keeps its normal
font). On the card the row shows `empty = F1`, and the history list shows `F1` for it.

* A slip with 1–2 empty boxes still goes in by itself. **3 or more empty boxes wait for you** (that usually means the camera
  missed something, e.g. a blank or badly lit slip). Change this with `auto_enter_max_blank`, or set `auto_enter_blanks` to `false` to
  always review slips that have an empty box.
* If you click **F1** on a row yourself, the same happens.
* **Undo last** removes the `F1` *and* puts the cell's previous font size back.
* Don't want the rule? Set `blank_marker` to `""` in `config.json` and empty boxes stay empty.

**`scans/`** keeps a picture of every entered slip plus `log.csv`, so any dispute can be checked against what the camera saw.

## 5. Reading the handwritten name

Handwriting is where ordinary OCR fails, so the name is read with Claude vision (`claude_model` in `config.json`,
default `claude-opus-5-5`, `claude_effort` low). Only **two small crops** (at most 800 px wide, sent as JPEG) - the printed title
strip and the NOMBRE strip - are sent, never the whole slip. Cost is a fraction of a cent per slip.

The reads run in a **background pool** (`name_workers`, default 6 at a time), never on the window's thread, so a slow answer
cannot hold up the stack. Each request waits at most 25 s (5 s to connect) with one retry, and the app gives up on a slip's
name `name_timeout_s` (45 s) after its request started (waiting for a free reader does not count) - it then writes `?` and puts
the slip on the review list. If names regularly queue up
("names being read" keeps growing in the strip - it then says "Claude is slower than the stack - raise `name_workers`") raise
`name_workers`; a name only waits for a free reader without being failed while other names keep coming back.

**No `ANTHROPIC_API_KEY`** / `read_names_with_claude: false` -> every slip opens the classic card and you type the name.
**The reader stops working in the middle of a stack** (key rejected, no credit, no internet, rate limit, no answer, an unexpected error inside the reader): after 3 such
failures in a row (or 3 requests silent for half a time-out) one red bar says **NAME READING IS NOT WORKING (reason)**, Auto-enter is
switched off so each slip waits on the card for you to type the name while it is in your hand, and there is no beep per slip. The picks
of the slips before that went in as usual, their names are `?` (red rows). The first name that is read correctly again clears the bar
and ticks Auto-enter again by itself (the switch-off is not saved).

## 6. Settings (`config.json`, created when you close the app)

`config.json` only contains the settings that differ from the defaults, plus `config_version`. It (and `.env`) may be saved by any
editor: UTF-8 with or without BOM, UTF-16 (what PowerShell's `>` writes) or the Windows ANSI code page. A `config.json` that cannot be
parsed is **never silently replaced**: it is kept as `config.json.bak` (`.bak2` ...), the window says so in the red bar and uses the
defaults; a value that cannot work (`matches_per_slip: 0`, a NaN time-out, `camera_index: -3`, `sheets: [1, 2]` ...) falls back to its
default with a note. **Upgrading:** a `config.json`
written by an older build (no `config_version`, or older than 2) saved *every* setting, which would pin the old slow defaults; so
`stable_frames`, `auto_enter_delay_s` and `auto_enter` are dropped when it is read (the window tells you) and everything else is
kept. Delete the file to start from the defaults.

| key | meaning |
|---|---|
| `camera_index`, `camera_backend` | which webcam; `dshow` (default, best for C920s) / `msmf` |
| `autofocus`, `manual_focus` | focus mode and the manual value (0–255; the Focus slider changes it) |
| `matches_per_slip` | rows on the slip (9) |
| `workbook` | blank = the workbook active in Excel at connect time, or part of a file name |
| `sheets` | sheet-name patterns to search (`["ENTRY*"]`) |
| `name_header` | header text above the names (`NOMBRE`) |
| `name_tag`, `name_tag_position`, `tag_separator` | remembered group-tag letters, `before` / `after`, and what goes between tag and name (default one space) |
| `uppercase_names` | write names in capitals |
| `auto_enter` | hands-free entry (default on) |
| `auto_enter_delay_s` | `0` (default) = picks go in at once, names follow; above 0 = old mode: wait for the name, then count down this many seconds |
| `stable_frames` | frames the slip must sit still before it is read (default 4) |
| `name_workers`, `name_timeout_s` | simultaneous name reads (6) and when to give up on one (45 s) |
| `unsure_marker` | appended to a name Claude was unsure of (`" ?"`; `""` = no marker, the row is still red) |
| `claude_model`, `claude_effort` | model and effort for the name read |
| `sounds` | beeps for "entered" / "needs you" (Windows only; default on) |
| `blank_marker`, `blank_font_size` | text and font size for an empty box (`"F1"`, `9`; `""` = leave empty) |
| `auto_enter_blanks`, `auto_enter_max_blank` | whether empty boxes may auto-enter, and how many (default yes, up to 2) |
| `save_after_write` | also press Save in Excel after each ticket (AutoSave normally covers this) |
| `expected_jornada` | blank = take it from the workbook name |
| `min_ink`, `multi_ratio` | mark sensitivity - run `run.bat scan photo.jpg` to see the ink readings per cell |

**`scanner.log`** (next to `config.json`) gets one line per entered slip: time, ticket, `read_ms` (reading the picture),
`excel_ms`, `name_ms`, status (`ok` / `review` / `wrong` / `undone`), the name written, plus how long the window's thread took
and the gap since the previous slip - and lines for start-up, Excel/camera problems, the breaker, `ALERT`s, `DUPLICATE_SKIPPED`,
`READER_DOWN` / `READER_OK`, `AUDIT_ERROR` and `CLOSE_DROPPED` (slips that were still on the card or in the queue when you closed the window anyway). If the app is slow or misbehaves on your PC, send that file.

**`scans/log.csv`** (UTF-8 with BOM, so Excel shows accents right) has one row per entered slip next to its picture. If the file cannot be
written (typically: it is open in Excel) the rows go to `scans/log-<time>.csv` instead, the red bar says so and nothing is lost; the same
happens for a full disk (the bar also warns below 1 GB free).

## 7. Command line

```
run.bat                          the scanner window
run.bat --demo slip.jpg          same, with a photo as the "camera" (it replays the same slip over and over, so use a copy of the workbook)
run.bat --demo slip.jpg --xlsx pool.xlsx   edit a closed .xlsx file instead of live Excel (this run only, not saved)
run.bat scan slip.jpg --out r.jpg [--name]   read one photo, print picks (+ name with Claude), save the annotated image (and r_name.jpg, the name picture)
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
| slips do not enter by themselves | **Auto-enter** must be ticked (a wrong-jornada slip unticks it), a name reader must exist (API key), and `auto_enter_delay_s` must be 0; an old `config.json` is fixed automatically (§6) |
| the next slip of the stack is not picked up | lift the previous slip completely off the pile, or press **Capture now**; "Same slip still there" means the one that was entered (and not undone) is still on top |
| names show `?` | open the red rows in the list and type them; `scanner.log` has the reason (no internet, timeout, ...). Many in a row = the red bar "NAME READING IS NOT WORKING" (§5): check the key / credit / internet |
| red bar "N slips are WAITING behind the card" | finish the open card (Enter / Esc): the waiting slips are entered right after, in order |
| red bar "AUDIT LOG PROBLEM" | `scans/log.csv` is open in Excel or the disk is full; rows go to `log-<time>.csv`, the slips are in your workbook anyway |
| the window did not fit my screen | it adapts to 1366x768 and up; the middle part scrolls if needed (mouse wheel) |
| it feels slow | look at `scanner.log` (`read_ms`, `excel_ms`, `name_ms`, `ui_ms`) or the strip under the preview |

## 9. What was tested, and what was not

Tested (automated - run `pytest`; the tests that use your real photo are skipped unless `QUINIELA_REAL_SLIP` points to it): the slip
reader on your real photo (all 9 marks, row for row) and on degraded versions of it; the capture engine on 30 fps sequences of whole
stacks (every slip read exactly once, none twice); the Excel fast path against a stand-in for Excel's object model (one small read +
one write per slip, no full re-read); and the real window driven end to end: a stack of 8 different slips played at 30 fps (1.1 s
per slip, and again at 0.63 s per slip) with name answers that take random seconds and arrive out of order - 8 rows in ticket order,
every name in its own row, the camera never paused, each slip's picks in Excel well under a second after it lay still; plus the tag
snapshot, review list, `?` marker, wrong-jornada breaker, classic card, undo while a name is in flight, config migration, and the name
reader's failure paths; and the failure cases found in review: several slips queued behind a card, closing the window with names in
flight, a failing name reader, a locked `log.csv`, config files in other encodings, small screens, keys typed in entries.

**Not tested on real hardware** - there was no Windows PC, webcam, Excel or API key in the build environment: the live **C920s**
(autofocus hunting, exposure changes after a hand passes, the focus slider), the **Excel COM connection** (verified only against a
stand-in; real Excel's speed, a cell in edit mode, `ScreenUpdating`), the **live Claude name read** (request shape verified against a
mock, latency unknown) and the **beeps**. Known blind spot: a slip swapped so fast that no hand or movement is ever seen (e.g. during the
~0.2 s it takes to read one) is not noticed - the app then simply does not enter the new slip; lift it again or press Capture now.
Do a first run on a *copy* of your workbook.

Developers: `pip install -r requirements-dev.txt`, then `pytest` (GUI tests need a display; on Linux use `xvfb-run pytest`).
