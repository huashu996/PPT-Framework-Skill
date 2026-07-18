# Academic Figure Fast Quality Gate

Use this checklist once after the authoritative PPT is assembled. Run it again only for checks that failed after the single targeted repair pass.

For the 5–8 minute single-page fast path, use `scripts/fast-ppt-audit.ps1` plus one finalized PowerPoint render. Do not also run `slides_test.py` or another full secondary audit unless the fast audit fails, cannot inspect the file, or the user explicitly requests high assurance.

## Blocking acceptance gate

Reject the figure when any of these remains:

- Missing, clipped, empty, unreadable, or out-of-bounds content.
- Ordinary text or a formula intersects an unrelated text box, image, frame, module, connector, arrowhead, divider, or panel boundary after applying the confirmed external gap token.
- A caption touches its image or enclosing boundary.
- A source-horizontal or source-vertical route is diagonal.
- An orthogonal route contains a non-horizontal/non-vertical segment.
- A connector reaches the wrong semantic object, a label, or a decorative fragment.
- An arrowhead is on the wrong end or points away from its destination.
- A required formula is missing, noneditable, duplicated, empty, or outside its reserved rectangle.
- A required native logical object was replaced by a full-figure raster.
- The saved final file cannot reopen.

Fix all blocking failures. Do not start another full inspection for minor decorative, crop, icon-style, or pixel-level spacing differences unless the user requested exact high assurance.

## One automated object audit

Run these checks together:

1. Compare expected and actual IDs/counts for modules, text, formulas, assets, routes, and outputs.
2. Confirm every object has positive dimensions and remains within the slide.
3. Detect text overflow and font substitution with the target PowerPoint renderer.
4. Inflate ordinary text/formula rectangles by the shared gap token and test them against unrelated occupied rectangles.
5. Confirm image aspect ratio and effective placed resolution.
6. Confirm peer modules use the intended shared dimensions, baselines, and gaps.
7. Confirm panel titles remain inside a dedicated title band.
8. Confirm requested files save, reopen, and are nonempty.

Exclude only documented intentional containment or overlay pairs. Do not globally ignore all text-versus-shape intersections merely because some text is hosted inside a module.

## Formula gate

Confirm the formula technology matches the contract.

For MathType:

- Compare the expected formula names and count with the saved deck.
- Require positive dimensions and slide bounds.
- When exposed, require `OLEFormat.ProgID = Equation.DSMT4`.
- Reopen the saved formula bank and merged PPT.
- Reopen the first, middle, and last equations through the PowerPoint MathType ribbon.
- Reopen every equation only when a sampled equation fails or the user explicitly requests high assurance.

Reject formula screenshots unless the user explicitly authorized a raster exception. Do not silently mix MathType, Office-native equations, and ordinary editable text.

## Connector gate

Use the final route JSON/TSV and frozen object geometry.

For every route:

- Confirm unique route ID, source, destination, declared anchors, arrowhead end, path kind, and expected visible arrow count.
- Confirm each endpoint lies on the intended semantic boundary anchor.
- Require `y1 == y2` for every horizontal segment and `x1 == x2` for every vertical segment.
- Require every segment of an orthogonal route to pass one of those axis tests.
- Require a straight route to contain exactly two waypoints.
- Confirm independent directed transitions and buses were not accidentally merged.
- Confirm ordinary span is at least 12 pt and the shaft remains visibly longer than the arrowhead.
- Confirm shafts and arrowhead envelopes do not intersect inflated text/formula/image rectangles, unrelated modules, or route-label corridors.
- Confirm curves remain native curves and one folded route remains one native elbow/polyline when PowerPoint supports it.

Generate connectors only after these numeric checks pass. Do not rely on automatic connection sites when they introduce a diagonal, hook, or wrong anchor.

## Editability gate

- Confirm logical boxes, text, formulas, labels, frames, nodes, arrows, and connectors are independently editable.
- Confirm local rasters are independently selectable, croppable, resizable, and replaceable.
- Reject a full-slide or full-figure screenshot.
- Confirm moving a named module either preserves connector attachment or triggers deterministic endpoint recomputation.

## One rendered-image inspection

Export one whole-slide preview after the automated audit.

Inspect:

- whole-page hierarchy, balance, margins, and orientation;
- text clarity, title separation, captions, and obvious collisions;
- arrow direction, endpoint contact, visible shaft length, and axis orientation;
- formula rendering and local raster quality;
- only audit-flagged or connector-dense regions at approximately 150%.

Permit one targeted same-PPT repair pass. Keep the finalization PowerPoint process open while fixing reported text, crop, route, and z-order failures; re-run only failed automated checks, overwrite the same preview, and stop when no blocking failure remains.

## High-assurance extension

Use only when explicitly requested:

- pixel-level comparison with the source;
- every-equation ribbon reopening;
- exhaustive 200–400% local inspection;
- per-object numeric symmetry reporting;
- repeated micro-adjustment beyond the single repair pass.

Do not activate this extension by default.
