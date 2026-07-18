# Portable Fast MathType PowerPoint Workflow

Use the visible MathType/WIRIS/Design Science plugin in the PowerPoint ribbon as the primary insertion and editing route. Resolve every helper relative to the selected `SKILL.md`. Never hard-code a MathType executable, username, drive letter, Codex home, or repository path.

## Choose the fast route

- For one to four formulas, insert directly into the authoritative PPT.
- For five or more formulas, assign one formula agent and create a separate `mathtype_formula_bank_test.pptx` while the main agent builds the non-formula layout.
- Give the formula agent exclusive ownership of desktop PowerPoint UI during the batch. Other agents must not automate the same PowerPoint session concurrently.

## Create the formula ledger once

Create `formula-targets.json` with one record per equation:

```json
{
  "name": "MATH_LOSS",
  "tex": "L=\\sum_i w_i e_i",
  "target_slide": 1,
  "reserved_rect_pt": [300, 180, 90, 24],
  "apparent_size_pt": 10
}
```

Freeze names and order after production starts. The main PPT uses matching named placeholders and does not wait for formula insertion.

## Run one ribbon smoke test

1. Open desktop PowerPoint once.
2. Select the MathType tab and invoke its new-equation command.
3. Enter one smoke-test formula and return to PowerPoint.
4. Confirm the inserted object is editable through the ribbon.
5. When exposed, require `OLEFormat.ProgID = Equation.DSMT4`.
6. Freeze this insertion recipe for the whole batch. Do not repeat add-in, registry, script, or path discovery for each equation.

Windows UI Automation may select the MathType tab and invoke the enabled visible insert control. Do not spend time attempting `PowerPoint.Application.Run` against private ribbon callbacks when the visible UI control works.

## Build one formula bank

Keep one PowerPoint process and one formula-bank presentation open.

For every ledger item:

1. Invoke the validated MathType ribbon insert command.
2. Enter the normalized TeX/content through the MathType editor.
3. Return to PowerPoint.
4. Rename the new equation to the exact ledger target name.
5. Place it in the next labeled grid cell.
6. Continue without final-size tuning.

Save at controlled checkpoints and after the final item, not after every equation. If one equation fails, mark that record and continue; repair failed records at the end.

## Merge once

After the bank and authoritative PPT are ready, open both in the same PowerPoint instance. For every target name:

1. Copy the matching bank equation.
2. Paste it on the target slide.
3. Preserve the exact name.
4. Scale proportionally to fit the reserved rectangle.
5. Center it horizontally and vertically.
6. Remove the placeholder only after a successful match.

Do not force width and height independently. Correct baseline alignment only after the full batch is placed.

## Validate the saved batch once

The default fast validation is:

- expected formula count equals the ledger count;
- every target name exists exactly once;
- every equation has positive dimensions and remains on its slide;
- each saved object is editable; when exposed, require `Equation.DSMT4`;
- the bank and merged PPT reopen successfully;
- the first, middle, and last formulas reopen through the MathType ribbon edit command.

Reopen every equation only when the user explicitly requests high assurance or a sampled equation fails.

## Optional helper interface

When `scripts/mathtype-ppt.ps1` exists, use it for one detection pass and one saved-object validation pass:

```powershell
$skillRoot = '<directory containing the selected SKILL.md>'
$tool = Join-Path $skillRoot 'scripts\mathtype-ppt.ps1'

& $tool -Action detect
& $tool -Action inspect -PptPath 'C:\path\figure.pptx'
& $tool -Action validate -PptPath 'C:\path\figure.pptx' -ExpectedCount 12
```

Use direct `Shapes.AddOLEObject(..., 'Equation.DSMT4')` only as a compatibility route when the ribbon cannot be operated and the ProgID is registered. The editor does not expose a universally reliable text-setting COM property, so do not pretend that ordinary PowerPoint text populates an equation.

If the helper is missing, create a temporary task-build adapter that:

- enumerates MathType/WIRIS/Design Science PowerPoint add-ins;
- enumerates equation shape names, positions, dimensions, and ProgIDs when exposed;
- validates expected count, unique names, positive dimensions, slide bounds, and successful reopen.

Use Office-native equations or grouped editable math text only when the user authorizes fallback. Never substitute a screenshot silently.
