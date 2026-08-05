# pptx-scenario-layout

**English:** this file · **Русский:** [README.md](README.md)

Update PowerPoint master presentations for corporate courses driven by a lesson
scenario. Proven on decks of 200–500 slides. Three pillars:

1. **Clone, don't create.** New content is never built from scratch — it's a
   copy of an existing template slide with the text replaced. Branding (frame,
   fonts, background) is preserved, layouts don't break.
2. **The scenario is the source of truth.** Order and content come from the
   scenario, even when the file disagrees.
3. **Reorder from end to start.** Slides are moved via PowerPoint COM
   (`MoveTo`) from the last position to the first — already placed slides are
   never shuffled by later moves.

Grown from a real case: 202 slides, 97 moves, ~40 minutes instead of 2 days of
manual shuffling. Validation: order matches the map, hidden slides preserved,
content intact.

## Why

When actualizing training courses, classroom-based and beyond: you need to add
new content, update old content, hide the obsolete and reorder everything to
follow the scenario. By hand that's hours of work, frayed nerves and mistakes.

## Full workflow

```
scenario (docx)  ->  operations (clone/edit/hide)  ->  order map  ->  validation
```

1. **Analyze the scenario.** New content is marked in the scenario (e.g. by
   color). Build a list of operations: what to add, change, hide. Do not put
   facilitator remarks (`Note:`, `Instructor:`) on slides.
2. **Operations** (`update.ps1`): clone a template slide + replace text +
   tag "New"/"Changed"; edit an existing slide; hide stale content. How to
   pick a template by content structure — see [TEMPLATES.md](TEMPLATES.md).
3. **Order map** (`layout.ps1`): final slide order = scenario map.
   Reorder from end to start.
4. **Validation** (`validate.ps1`): order = map, hidden slides preserved,
   content in place.

## Requirements

- Windows + PowerPoint installed (COM).
- PowerShell 5.1+.

## Usage

### Step 1. Operations: add / change / hide

Create `operations.json` (example: [operations.example.json](operations.example.json)):

```json
{
  "operations": [
    { "op": "edit",  "find": "Beta task",        "replace": { "Beta task": "Beta task v2" }, "tag": "Edited" },
    { "op": "clone", "from": 2, "to": 4,          "replace": { "Beta theory": "New theory" }, "tag": "New" },
    { "op": "hide",  "find": "Omega outro" }
  ]
}
```

- `edit` — find a slide by text (`find`) or by number (`slide`), replace
  substrings in shapes, add a tag.
- `clone` — duplicate the template slide `from`, replace text, add a tag,
  move it to the target position `to`.
- `hide` — hide the slide (`show="0"`); it stays in the file but is not shown.

Run (creates a `_backup.pptx` backup copy):

```powershell
powershell -ExecutionPolicy Bypass -File update.ps1 -PptxPath ".\deck.pptx" -Operations "operations.json" -Backup
```

### Step 2. Order map and reorder

`target_order.txt`: line N = original slide index that must end up at
position N (1-based; empty lines and `#` are ignored). It must be a full
permutation: 0 gaps, 0 duplicates — otherwise the script won't run.

```powershell
powershell -ExecutionPolicy Bypass -File layout.ps1 -PptxPath ".\deck.pptx" -MapFile "target_order.txt" -Backup
```

### Step 3. Validate

```powershell
powershell -ExecutionPolicy Bypass -File validate.ps1 -PptxPath ".\deck.pptx" -MapFile "target_order.txt" -OriginalPath ".\deck_original.pptx"
```

`validate.ps1` checks:
- slide count matches the map;
- the `sldIdLst` order is sequential;
- the number of hidden slides (`show="0"`);
- if `-OriginalPath` is given — each slide's content matches the original
  (by normalized text; pure numbers — auto slide numbers — are ignored).

## Known issues (tested in practice)

- **Script encoding.** PowerShell 5.1 breaks `.ps1` with Cyrillic comments if
  the file is not saved with a BOM. Add Russian text only with UTF-8 BOM.
- **File names.** Paths may contain an em dash «—», not a hyphen. Search files
  by glob/`Where-Object`, don't hardcode strings.
- **Console.** Cyrillic in terminal output shows as garbage (cp866/cp1251).
  Write results to UTF-8 files and read those, don't rely on the console.
- **Renumbering.** After a COM-Save PowerPoint renames slide files:
  `slide{i}.xml` is the new i-th slide. Validate by text, not by old file
  numbers.
- **Finding a slide by text.** Use short unique substrings that don't cross a
  line break (`\x0b` in PowerPoint text breaks search).
- **Unused slides** — hide them (`show="0"`), don't move them to the end,
  otherwise it's hard for a person to navigate the file.
- **"Dirty" template.** After cloning, text boxes with the original text may
  remain — replace/clean them.

## Structure

```
update.ps1             # operations: clone / edit / hide (new content)
operations.example.json# example operations table
layout.ps1             # reorder by map (MoveTo, end to start)
validate.ps1           # post-update checks
target_order.txt       # example order map
TEMPLATES.md           # how to pick a template by content structure
```

## License

MIT.