# Portable Fast MathType PowerPoint Workflow

Use the visible MathType/WIRIS/Design Science plugin in the PowerPoint ribbon as the primary insertion and editing route. Resolve every helper relative to the selected `SKILL.md`. Never hard-code a MathType executable, username, drive letter, Codex home, or repository path.

## One-pass fast loop for every equation

Run this exact loop for each equation:

1. Prepare the equation content, target object name and reserved rectangle before touching the UI.
2. Open the authoritative PPTX by passing its absolute path directly to PowerPoint, or reuse the already-open title-matching window; then restore, activate and maximize it. Never use screenshots, the Open dialog, path typing or directory browsing to open the file.
3. Invoke the enabled MathType new-equation control by its current accessible name. Do not use screen coordinates when an accessible control exists.
4. In the empty editor, press `x` and inspect the visible editor state. Pass only when there is exactly one English `x` and no Chinese candidate bar, pinyin pre-edit text, composition underline or pending selection. An English `x` accompanied by a Chinese candidate bar is a failed test. Cancel and clear, switch to English, and test once more.
5. Clear the successful `x` test and enter the equation per key with the already validated MathType shortcuts.
6. Inspect the visible equation once, press `Ctrl+S`, close and return to PowerPoint.
7. Rename and proportionally place the newest nonempty `Equation.DSMT4`, delete its matching invisible placeholder, then save the PPTX.

Do not perform window enumeration, add-in detection, coordinate calibration, browser navigation, clipboard experiments or full-slide rendering inside this loop unless the current step has a concrete failure. A simple equation should normally complete in 30–90 seconds. If the route is still unstable after about two minutes, stop and report the exact blocker instead of looping.

## Non-negotiable direct-write route

Use this single route whenever the visible PowerPoint MathType tab is available:

1. Finish the non-formula layout in the one authoritative PPTX. Reserve one invisible, uniquely named rectangle per formula; do not create visible text equations.
2. Open that exact PPTX directly in desktop PowerPoint by passing the known absolute PPTX path to PowerPoint or the operating-system file association. If it is already open, activate the matching window instead. Never open blank PowerPoint first, never use `File > Open` or a system file picker, and never use screenshots, coordinates, typed paths, clipboard paths or directory navigation to locate the PPTX. Do not open a browser, help page, web MathType control, intermediate PPT, or formula bank for one to four formulas.
3. Treat the visible `MathType` tab and the enabled new-equation button inside the insert-equation group as authoritative. A detection helper that reports no add-in does not override a visible working ribbon control.
4. Invoke the new-equation button by its current accessible name or current PowerPoint KeyTip sequence. Never reuse a stale coordinate or element index.
5. Before entering every real equation, first activate and maximize the authoritative PowerPoint window, confirm its title, open a new blank MathType editor through the current visible control, and enter one Latin test character `x`. Only when the editor visibly shows a single English `x` with no Chinese character, Chinese candidate bar, pinyin pre-edit string, composition underline or pending selection may the test character be cleared and the real equation entered. A Latin `x` with a Chinese candidate bar still means the IME test failed. If any such state appears, cancel the candidate, clear the editor once, toggle to English once, and repeat the test once. If the second result cannot be visibly confirmed, stop; never infer success from the UI language, a sent key, the previous formula or the appearance of `x` alone.
6. Enter the equation directly in MathType. For classic MathType, prefer deterministic per-key input and native shortcuts after the smoke test succeeds, such as `Ctrl+L` for a subscript and `Ctrl+G` followed by a letter for a Greek symbol. Do not prewrite the formula in PowerPoint.
7. If clipboard or TeX paste fails once, stop using paste for the task. Do not alternate among clipboard, TeX, Unicode injection, coordinates, and different editor routes.
8. Press `Ctrl+S` in MathType to update the current OLE, then close and return to PowerPoint. Save the PowerPoint presentation after the object returns.
9. Identify the newest nonempty `Equation.DSMT4` object. Delete only the matching invisible placeholder, rename the OLE to the target ID, and position it once in the reserved rectangle. Scale proportionally and center it; do not stretch width and height independently.
10. Use COM or a helper only for enumeration, renaming, proportional placement, bounds checks, and saving. Never use ordinary PowerPoint text as formula content.
11. Repeat the same route for the remaining equations, including the per-equation `x` input-method test. Do not reopen MathType merely to adjust position.
12. After all formulas and layout edits are saved, reopen each of one to four formulas; for five or more, reopen the first, middle, and last. Confirm the content is nonempty and correct, then render the final saved PPTX and verify every equation is visible.

The only allowed stable fallback after one known-good OLE exists is to duplicate that verified `Equation.DSMT4`, rename and place the duplicate, open it through its OLE edit verb, replace its content, save, and use this same route for all remaining formulas.

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

## Run the ribbon and input test

1. Open the authoritative PPTX once by its absolute path, or reuse its already-open title-matching PowerPoint window. Do not open the Open dialog or take screenshots to locate and type the path.
2. Select the MathType tab and invoke the new-equation command whose accessible parent is the insert-equation group. Require an exact accessible-name match; never use a stale coordinate/index or a MathType help/web control.
3. Test the input state with `x` before every new equation. Require exactly one English `x` and no candidate bar, pinyin pre-edit text, composition underline or pending selection. If any IME candidate or Chinese state appears—even when `x` itself is Latin—cancel the candidate, clear once, toggle to English once, and retest once. Do not begin the real equation until the visible test passes; if it cannot be confirmed, stop and report the input-method blocker.
4. Enter one smoke-test formula and return to PowerPoint. If classic MathType rejects ordinary clipboard text or reports that the clipboard contains no equation data, stop retrying paste and use per-key keyboard input for this task.
5. Confirm the inserted object is editable through the ribbon.
6. When exposed, require `OLEFormat.ProgID = Equation.DSMT4`.
7. Confirm the update and return recipe once. For classic MathType, run `Ctrl+S` to update the current OLE, then use `File > Close and return`; accept the save confirmation dialog only when it still appears.
8. Reopen the smoke-test object and confirm the expected formula content is nonempty. Object existence and `ProgID` alone do not prove that content was saved.
9. Freeze the insertion entry, per-key input method, update command and return recipe for the whole batch. Recheck the actual input-method state with `x` for every equation because the IME may switch between editor sessions. Do not switch between ribbon, coordinates, COM helpers and clipboard experiments after a route has succeeded.

If the ribbon becomes unstable after one valid `Equation.DSMT4` object exists, the stable fallback is to duplicate that known-good OLE, rename and place the duplicate, then open the duplicate through its OLE edit verb and replace its content. Freeze this direct-object route for the remainder of the batch; do not alternate back to ribbon controls.

Windows UI Automation may select the MathType tab and invoke the enabled visible insert control. Do not spend time attempting `PowerPoint.Application.Run` against private ribbon callbacks when the visible UI control works.

## Build one formula bank

Keep one PowerPoint process and one formula-bank presentation open.

For every ledger item:

1. Invoke the validated MathType ribbon insert command.
2. Enter the normalized TeX/content through the MathType editor using the validated input method. When the smoke test required per-key input, do not switch back to clipboard paste during the batch.
3. Update the current OLE using the validated command, then return to PowerPoint using the validated close-and-save sequence.
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
6. After a successful match, delete every non-OLE placeholder carrying the same target name before rendering or saving the final file.
7. Confirm that the target name resolves to exactly one final equation object and that no placeholder text remains underneath or above it.

Do not force width and height independently. Correct baseline alignment only after the full batch is placed.

## Validate the saved batch once

The default fast validation is:

- expected formula count equals the ledger count;
- every target name exists exactly once;
- every equation has positive dimensions and remains on its slide;
- each saved object is editable; when exposed, require `Equation.DSMT4`;
- the final PowerPoint rendering shows nonempty formula content at every target; an empty OLE preview is a failure;
- the bank and merged PPT reopen successfully;
- for one to four formulas, every formula reopens with the expected content; for five or more, the first, middle, and last formulas reopen with the expected content.

Reopen every equation only when the user explicitly requests high assurance or a sampled equation fails.

## Optional helper interface

When `scripts/mathtype-ppt.ps1` exists, use it only for one detection pass and one saved-object inspection/validation pass:

```powershell
$skillRoot = '<directory containing the selected SKILL.md>'
$tool = Join-Path $skillRoot 'scripts\mathtype-ppt.ps1'

& $tool -Action detect
& $tool -Action inspect -PptPath 'C:\path\figure.pptx'
& $tool -Action validate -PptPath 'C:\path\figure.pptx' -ExpectedCount 12
```

Do not use the helper's `insert` or `edit` actions to produce equations. If any helper action raises a null-object error, conflicts with the active PowerPoint presentation, or creates an empty OLE, stop that helper branch immediately and do not retry it. Continue with the visible ribbon route already defined above.

Use direct `Shapes.AddOLEObject(..., 'Equation.DSMT4')` only as a one-time compatibility test when no visible ribbon control exists and the ProgID is registered. Never switch to it merely because detection failed while the ribbon is visible. The editor does not expose a universally reliable text-setting COM property, so do not pretend that ordinary PowerPoint text populates an equation.

If the helper is missing, create a temporary task-build adapter that:

- enumerates MathType/WIRIS/Design Science PowerPoint add-ins;
- enumerates equation shape names, positions, dimensions, and ProgIDs when exposed;
- validates expected count, unique names, positive dimensions, slide bounds, and successful reopen.

Use Office-native equations or grouped editable math text only when the user authorizes fallback. Never substitute a screenshot silently.

## Stop conditions that prevent looping

- Do not use screenshots, coordinates, typed paths, the Open dialog or directory browsing to open a known PPTX path; direct-path launch or reuse of the already-open matching window is mandatory.
- Do not enter any real formula character until the current editor has passed the visible English `x` test with no Chinese candidate or pre-edit state.
- Do not retry the same failed insertion, paste, helper, or input-method action more than once.
- Do not open a browser or MathType web/help control to recover a desktop equation workflow.
- Do not switch routes after one formula has been successfully inserted, saved, reopened, and verified.
- Do not reopen an equation for geometry changes; position and scale the OLE in PowerPoint.
- Do not perform repeated full-slide renders during formula entry. Render once after all formulas and layout changes are saved.
