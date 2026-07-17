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

1. Detect MathType before constructing formulas.
2. Reserve each formula's occupied rectangle during joint layout.
3. Insert one named MathType OLE object into each reserved rectangle.
4. Open the named object with the OLE `Edit` verb.
5. In the MathType editor, select the equation body, enter or paste the intended formula, and return to PowerPoint so the embedded object updates.
6. Save the PowerPoint.
7. Reopen the saved PowerPoint and inspect or validate the MathType objects.

Run commands from PowerShell:

```powershell
$tool = 'C:\Users\18144\.codex\skills\paper-fig-skill\scripts\mathtype-ppt.ps1'

& $tool -Action detect

& $tool -Action insert `
  -PptPath 'C:\path\figure.pptx' `
  -SlideNumber 1 `
  -ShapeName 'MATH_Loss' `
  -Left 300 -Top 200 -Width 90 -Height 24

& $tool -Action edit `
  -PptPath 'C:\path\figure.pptx' `
  -SlideNumber 1 `
  -ShapeName 'MATH_Loss'

& $tool -Action inspect -PptPath 'C:\path\figure.pptx'

& $tool -Action validate `
  -PptPath 'C:\path\figure.pptx' `
  -ExpectedCount 1
```

PowerPoint coordinates and sizes are in points. Use stable `MATH_*` shape names. `insert` creates a genuine editable MathType OLE object; it does not create a screenshot. `edit` opens the selected object through its MathType `Edit` verb and intentionally leaves the presentation open for interactive or UI-automated formula entry.

The legacy MathType OLE server does not expose a reliable text-setting COM property. Do not pretend that assigning a PowerPoint text string populates the equation. Enter formula content through the opened MathType editor, then use `validate` to verify the saved object type, count, dimensions, and slide bounds.

## Formula-writing lifecycle

For every formula, use this exact lifecycle:

1. Assign a stable semantic name such as `MATH_FinalAttention` and reserve its final layout rectangle.
2. Call `insert` once. Never replace the object with an image after it exists.
3. Call `edit` for that name. Use the visible MathType editor for manual or UI-automated content entry.
4. Return focus to PowerPoint and confirm the rendered formula changed inside the same named object.
5. Resize or reposition the existing OLE shape only after the formula content is final; preserve its aspect ratio unless the reference requires a controlled adjustment.
6. Save, close, reopen, and call `inspect`.
7. Call `validate` with the expected total object count. Also validate important formula names individually when omissions would be hard to notice.

Reject the result if copying from MathType produced only an enhanced metafile, picture, or ordinary text object. A formula is accepted as MathType-editable only when reopening the PPT still reports `OLEFormat.ProgID = Equation.DSMT4`.
