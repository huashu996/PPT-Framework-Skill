# Portable MathType PowerPoint workflow

Use the MathType plugin in the PowerPoint ribbon when the confirmed formula contract selects MathType. This is the primary insertion and editing route. Use `scripts/mathtype-ppt.ps1` for capability reporting, saved-object inspection, validation, and direct-OLE compatibility only.

Resolve every helper relative to the selected `SKILL.md`. A missing helper is an incomplete-package warning, not evidence that MathType itself is absent.

## Primary PowerPoint ribbon workflow

1. Open the authoritative PPT in desktop PowerPoint.
2. Navigate to the target slide and select the reserved formula location or an existing equation.
3. Enter through the MathType tab in the PowerPoint ribbon and choose the appropriate new-equation or edit-equation command.
4. Enter the formula in the MathType editor, then return to PowerPoint.
5. Place the equation inside its reserved rectangle and assign a stable `MATH_*` name when object control permits.
6. Save, close, and reopen the deck.
7. Select the equation and use the MathType ribbon edit command again. This round-trip is the definitive editability check.

Use desktop UI automation when unattended entry is required. Do not launch a machine-specific `MathType.exe` path and do not substitute a screenshot.

## Capability and compatibility interface

Detect the PowerPoint plugin by enumerating `PowerPoint.Application.COMAddIns` and `PowerPoint.Application.AddIns`. Match MathType, Design Science, or WIRIS in the add-in name, description, path, or ProgID, and record its connected or loaded state. If enumeration is inconclusive, visually check the PowerPoint ribbon.

MathType ribbon insertions commonly expose this legacy editable OLE interface. Use it for compatibility and validation when present, but do not require it before checking the ribbon plugin:

- OLE ProgID: `Equation.DSMT4`
- Registered description: `MathType 7.0 Equation`
- CLSID: `{0002CE03-0000-0000-C000-000000000046}`
- Compatibility insertion: `Shapes.AddOLEObject(left, top, width, height, 'Equation.DSMT4')`
- Compatibility editing: `shape.OLEFormat.DoVerb(2)`
- Saved-object check: `shape.OLEFormat.ProgID -eq 'Equation.DSMT4'`

Do not hard-code an executable path, username, drive letter, Codex home, or repository checkout. If the bundled helper is missing, create the temporary build-directory adapter defined in the main `SKILL.md`.

## Portable helper invocation

Replace the placeholder with the actual directory containing the selected `SKILL.md`:

```powershell
$skillRoot = '<directory containing the selected SKILL.md>'
$tool = Join-Path $skillRoot 'scripts\mathtype-ppt.ps1'

& $tool -Action detect
& $tool -Action inspect -PptPath 'C:\path\figure.pptx'
& $tool -Action validate -PptPath 'C:\path\figure.pptx' -ExpectedCount 1
```

Use `-Action insert` and `-Action edit` only as direct-OLE compatibility routes when the PowerPoint ribbon cannot be operated and `Equation.DSMT4` is registered:

```powershell
& $tool -Action insert `
  -PptPath 'C:\path\figure.pptx' `
  -SlideNumber 1 `
  -ShapeName 'MATH_Loss' `
  -Left 300 -Top 200 -Width 90 -Height 24

& $tool -Action edit `
  -PptPath 'C:\path\figure.pptx' `
  -SlideNumber 1 `
  -ShapeName 'MATH_Loss'
```

The MathType editor does not expose a universally reliable text-setting COM property. Do not pretend that assigning ordinary PowerPoint text populates the equation. Enter formula content through the editor opened from the PowerPoint ribbon.

Reject the result if insertion produced only an enhanced metafile, picture, or ordinary text object. Accept a formula as MathType-editable only when it reopens successfully through the MathType ribbon; when an OLE ProgID is exposed, also require `OLEFormat.ProgID = Equation.DSMT4`.
