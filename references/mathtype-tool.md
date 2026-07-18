# MathType PowerPoint tool

Use `scripts/mathtype-ppt.ps1` when the confirmed formula contract selects MathType. The tool links PowerPoint directly to the installed MathType OLE server and uses the registered `Equation.DSMT4` object type.

## Fixed integration interface

Treat this interface as authoritative. Do not search for the MathType interface again unless `-Action detect` reports that registration is missing:

- OLE ProgID: `Equation.DSMT4`
- Registered description: `MathType 7.0 Equation`
- CLSID: `{0002CE03-0000-0000-C000-000000000046}`
- PowerPoint insertion call: `Shapes.AddOLEObject(left, top, width, height, 'Equation.DSMT4')`
- PowerPoint editing call: `shape.OLEFormat.DoVerb(2)` where verb 2 is `Edit`
- Required saved-object check: `shape.OLEFormat.ProgID -eq 'Equation.DSMT4'`

The bundled tool resolves the MathType executable from the OLE registration. Do not hard-code or rediscover the executable path in ordinary figure work.

## Required sequence

1. Run `preflight` before constructing the figure. It validates MathType and obtains a usable PowerPoint application, preferring an already active instance and otherwise creating a private instance.
2. Reserve each formula's occupied rectangle during joint layout.
3. Insert one named MathType OLE object into each reserved rectangle.
4. Open the named object with the OLE `Edit` verb.
5. In the MathType editor, focus the equation body, select all, press Backspace, then enter the intended formula. Save with `Ctrl+S` and close the editor with `Ctrl+F4` so the embedded object updates in PowerPoint.
6. Call `fit` with the reserved rectangle. MathType commonly expands the OLE object to its intrinsic equation size after editing; `fit` preserves the equation aspect ratio, scales it inside the reserved rectangle, and centers the same object.
7. Repeat edit and fit for every formula while reusing the same active PowerPoint presentation. Save and close once after the full batch.
8. Reopen the saved PowerPoint once and run `inspect-validate` to inspect and validate all MathType objects in one process.

Run commands from PowerShell:

```powershell
$tool = 'C:\Users\18144\.codex\skills\paper-fig-skill\scripts\mathtype-ppt.ps1'

& $tool -Action preflight

& $tool -Action insert `
  -PptPath 'C:\path\figure.pptx' `
  -SlideNumber 1 `
  -ShapeName 'MATH_Loss' `
  -Left 300 -Top 200 -Width 90 -Height 24

& $tool -Action edit `
  -PptPath 'C:\path\figure.pptx' `
  -SlideNumber 1 `
  -ShapeName 'MATH_Loss'

& $tool -Action fit `
  -PptPath 'C:\path\figure.pptx' `
  -SlideNumber 1 `
  -ShapeName 'MATH_Loss' `
  -Left 300 -Top 200 -Width 420 -Height 48

& $tool -Action inspect-validate `
  -PptPath 'C:\path\figure.pptx' `
  -ExpectedCount 1
```

PowerPoint coordinates and sizes are in points. Use stable `MATH_*` shape names. `insert` creates a genuine editable MathType OLE object; it does not create a screenshot. `edit` opens the selected object through its MathType `Edit` verb and intentionally leaves the presentation open for interactive or UI-automated formula entry.

The legacy MathType OLE server does not expose a reliable text-setting COM property. Do not pretend that assigning a PowerPoint text string populates the equation. Enter formula content through the opened MathType editor, then use `inspect-validate` to verify the saved object type, count, dimensions, and slide bounds in one call.

The tool never closes unrelated user presentations. It first tries the active PowerPoint application because some Office installations return an unusable `New-Object -ComObject PowerPoint.Application` wrapper with a null `Presentations` collection. When no usable active instance exists, the tool creates and owns a private PowerPoint instance. If preflight asks for PowerPoint to be opened once, open it and rerun instead of rediscovering COM registration or switching to an improvised insertion method.

## Safe UI automation sequence

Use this sequence for each named formula object:

1. Open the object with `edit` and wait for a visible MathType window.
2. Refresh the current window state and focus the equation canvas.
3. Press `Ctrl+A`, then Backspace. Do not type or paste while the prior equation remains selected.
4. Enter the replacement equation.
5. Press `Ctrl+S`, wait for the update, then press `Ctrl+F4`.
6. Wait until the MathType window closes before calling `fit` or opening the next object.

If MathType reports that the clipboard does not contain equation data, dismiss the dialog, close the editor without accepting an empty replacement, reopen the same named object, and repeat with the delete-before-input sequence. Do not count the object as edited merely because the automation loop advanced.

Avoid multiline equations in narrow portrait rectangles. `fit` preserves aspect ratio; a tall, narrow equation can become a hairline-width or black-bar artifact. Match the reserved rectangle to the natural equation aspect ratio and inspect the rendered result. Keep a nonformula variable inventory in adjacent native PowerPoint text only when the confirmed formula contract permits that separation; keep every actual formula in MathType.

## Formula-writing lifecycle

For every formula, use this exact lifecycle:

1. Assign a stable semantic name such as `MATH_FinalAttention` and reserve its final layout rectangle.
2. Call `insert` once. Never replace the object with an image after it exists.
3. Call `edit` for that name. Use the visible MathType editor for manual or UI-automated content entry.
4. Save with `Ctrl+S`, close the editor with `Ctrl+F4`, return focus to PowerPoint, and confirm the rendered formula changed inside the same named object.
5. Call `fit` once with the reserved formula rectangle. Do not manually guess the post-edit size or replace the expanded object.
6. Repeat steps 1–5 for every formula in the same active PowerPoint/MathType session.
7. Save once, close once, reopen once, and call `inspect-validate` with the expected total object count. Also validate important formula names individually when omissions would be hard to notice.

Reject the result if copying from MathType produced only an enhanced metafile, picture, or ordinary text object. A formula is accepted as MathType-editable only when reopening the PPT still reports `OLEFormat.ProgID = Equation.DSMT4`.
