---
name: paper-fig-skill
description: Create publication-ready, layout-first and symmetry-aware academic framework, technical-route, chapter-overview, and research-content figures from papers, folders, outlines, formulas, palettes, and reference figures. Use when Codex must analyze research materials, control content density and typography, reproduce or optimize a reference layout, segment complex source imagery into controlled raster components, and deliver an editable PowerPoint, high-resolution PNG, and Word figure version for a thesis, journal manuscript, or presentation.
---

# Paper-Fig-Skill

Turn research evidence and framework materials into a complete academic figure. Preserve the research content while improving hierarchy, visual balance, and editability.

## Run the always-confirm intake gate

At the start of every invocation, send one compact confirmation message. Prefill known values, but explicitly ask the user to confirm these decisions every time because they control the entire layout:

1. Intended use and exact master/canvas: page or slide size, orientation, usable margins, and any required aspect ratio.
2. Typography contract: Chinese and Latin fonts, permitted size range, exact number of size levels, and whether mixed-script runs may use different fonts.
3. Formula contract: whether formulas must use editable MathType OLE objects, Office-native editable equations, or grouped editable text; confirm the formula size range separately when it differs from ordinary text.
4. Content-editing boundary: preserve wording verbatim or allow shortening, rephrasing, line balancing, and removal of decorative icons.
5. Visual contract: exact reference reproduction or optimized adaptation, color/style preference, and whether apparent symmetry and orientation must be preserved.
6. Output contract: editable PPT, high-resolution PNG, Word version, and any explicitly allowed rasterized exceptions.

Also ask for missing materials, the target framework/chapter/study, reference figures, and output language. A value stated in the current request may be presented as a proposed confirmed value, but still ask for one concise confirmation before drawing. Do not start construction until these controlling choices are confirmed. If no Chinese font is supplied, propose 宋体 rather than silently assuming it.

## Follow the phase-gated layout-first workflow

Execute these stages in order. Treat every stage as a hard gate: do not start a later stage until the current stage has a recorded, internally consistent result. Do not fill provisional boxes with text and stretch the finished diagram afterward.

1. **Confirm the generation contract.** Confirm the master, usable margins, fonts, ordinary text size range, formula technology, formula size range, content-editing boundary, visual fidelity, and outputs. Stop before construction if MathType use or either size range is unknown.
2. **Understand the complete system and route logic.** Read the supplied images and text together. Identify the central problem, modules, submodules, hierarchy, inputs, outputs, semantic peer groups, and reading order. Create a route ledger before drawing: for every visible or implied flow, record the source, destination, direction, arrowhead end, straight/elbow/curve/branch type, junction behavior, anchor side, line weight, solid/dashed style, and z-order. Distinguish a shared bus, a merge node, and several independent arrows. Do not infer a connection from visual proximity. Do not proceed until every source and destination is accounted for.
3. **Lock content geometry before container geometry.** Set the confirmed PPT master and usable figure bounds. Lock the exact wording, formulas, fonts, size levels, wrapping, symbols, pictures, and icons. Render or probe the final text and formulas at their real sizes; decide their occupied rectangles and intended positions. Add internal padding and only then derive text boxes, formula boxes, module boxes, system panels, and the complete figure bounds. Never fix a box first and shrink or crowd its content afterward.
4. **Classify each whole visual object.** Inventory every non-text element as one of: native editable PowerPoint shape/connector, editable equation object, tightly cropped source raster, or newly generated local raster. Prefer one native PowerPoint primitive for a simple semantic object. Preserve the aspect ratio and safe crop of raster assets. Record stacking before construction. Split an object only for a true branch, independently attached endpoints, or geometry no single native object can express.
5. **Lay out all content together without connectors.** Treat text, formulas, shapes, pictures, symbols, titles, labels, arrowheads, and reserved connector corridors as simultaneous occupied geometry. Place the complete composition, not one object category at a time. Detect symmetry, mirror paired coordinates, normalize peer sizes, distribute whitespace, and reserve title bands and route-label corridors. Reject the layout if any unrelated occupied rectangles overlap, touch a boundary, have visually sticky gaps, or leave avoidable local crowding while usable space remains elsewhere. Do not draw logical arrows during this stage.
6. **Add connectors only after the layout passes.** Re-read the route ledger and connect the already placed modules from their intended boundary anchors. Confirm direction first, then choose the native path type. Prefer horizontal or vertical straight routes, then balanced orthogonal folds, then meaningful diagonals. Use a true PowerPoint curve for a curved source path and one native elbow/polyline for one folded route. Reserve visible shaft length and use a filled or open arrowhead only as confirmed by the source. Reject any route that crosses text, a formula, a picture, an unrelated module, another arrowhead, or a reserved label corridor.
7. **Converge by editing the same PowerPoint.** Freeze the first complete editable draft as the authoritative artifact. Render it, inspect it at whole-page and 100–200% local views, and move, resize, crop, reorder, or reroute the existing objects. Recompute affected connectors after moving modules. Do not regenerate for ordinary spacing, overlap, line routing, z-order, title, formula, or label corrections. Repeat same-file micro-adjustment and QA until the acceptance checklist passes.

Maintain three explicit working records through the workflow: the route ledger, the object/asset inventory, and the occupied-rectangle layout map. If a later correction invalidates one record, update it before changing the PowerPoint.

## Use the bundled MathType tool

When the confirmed formula contract selects MathType, read [references/mathtype-tool.md](references/mathtype-tool.md) and call [scripts/mathtype-ppt.ps1](scripts/mathtype-ppt.ps1). Do not rely on an unverified manual paste or a formula screenshot.

Treat the bundled script and its documented `Equation.DSMT4` interface as the permanent MathType connection for this skill. Do not search the registry, rediscover the ProgID, or invent a new insertion method during ordinary figure production. Investigate the installation only when the tool's `-Action detect` fails.

Use the tool in this order:

1. Run `-Action detect` before figure construction and stop if the `Equation.DSMT4` OLE server is unavailable.
2. Reserve and name every formula rectangle during the joint occupancy stage.
3. Run `-Action insert` to create one genuine editable MathType OLE object per formula at the reserved PowerPoint coordinates.
4. Run `-Action edit` for the named object and enter the formula through MathType. Use UI automation when unattended entry is required; do not assign ordinary PowerPoint text to an OLE equation.
5. Return to PowerPoint so MathType updates the same embedded object, then save and close the PPT.
6. Reopen the PPT, then run `-Action inspect` and `-Action validate`. Reject missing, substituted, off-slide, or non-MathType formula objects.

Keep stable `MATH_*` object names so later size, position, and collision adjustments target the same formula objects in the authoritative PPT.

## Apply the hard layout-priority contract

When fidelity conflicts with readability, apply this order without exception:

1. Eliminate accidental overlap and occlusion.
2. Preserve strong, repeatable whitespace.
3. Balance and align the complete composition.
4. Preserve reference positions and local similarity.

Treat text, formulas, frames, images, nodes, symbols, arrow shafts, arrowheads, connector labels, titles, panel boundaries, legends, and annotations as occupied geometry. A layout is not valid merely because each object fits inside its own bounding box. Measure and inspect clear edge-to-edge distance between unrelated occupied regions.

Allow overlap only when it is a confirmed semantic construction in the source, such as an editable variable or symbol intentionally placed over a dataset image, or a node intentionally covering a continuous lower-layer path. For an intentional image overlay:

- keep the overlay entirely inside the image's safe interior;
- retain visible clearance from the image edge and every enclosing frame;
- place the overlay above the image and verify legibility;
- do not use the exception to justify text touching a border, arrow, unrelated icon, or neighboring module.

Reserve a dedicated title band inside every colored panel or major container. The title band must not share occupied space with the first row of nodes, boxes, formulas, labels, or connectors. Give the title balanced clearance to the panel boundary and to the first content row; move the content group, not only the title, when the band is too tight.

Treat labels on or near a line as a layout error unless the source explicitly uses the line as mathematical notation. Place connector labels, route names, and annotations above or below the route in a reserved label corridor. Keep the label clear of the line stroke, arrowhead, module border, and neighboring object. Text boxes must use no fill and no outline by default; use a filled or outlined text container only when the reference depicts a real labeled shape.

Place ordinary labels and formulas on the top z-order, but do not use z-order to hide a collision. Reflow the geometry first. Use z-order only for intentional semantic stacking: container at back, complete path next, overlay node or image next, and label at top.

Use this mandatory construction order inside every module group:

1. Lock the wording, font, size level, wrapping rule, symbols, icons, images, formulas, and permitted manual breaks.
2. Measure the rendered bounds of every text, formula, symbol, node, image, arrowhead, and connector label together.
3. Reserve non-overlapping occupancy rectangles for the complete group, including padding, title bands, symbol-to-label gaps, image clearances, formula clearances, and connector channels.
4. Fit the complete peer group inside its parent and redistribute available whitespace.
5. Create the parent, text, formula, graphic, and peer objects at the computed coordinates as one coordinated operation.
6. Verify the complete group before adding connections; do not treat any one object category as a later afterthought.
7. Add each continuous path as one native connector, line, curve, or polyline whenever possible; send it behind any overlay nodes. Split it only at a true branch or independently attached endpoint.

Do not treat “text fits inside its own text box” as a completed layout check. Text can pass overflow tests and still cover a neighboring node, border, arrow, or label.

Treat the construction order as a hard gate: understand the logic, fix the master and complete figure bounds, lock typography, lock wording, measure final text, derive box sizes, reserve all block rectangles, place the layout, and only then add connectors. Never create fixed boxes first and force text into them afterward.

After the first complete render, freeze the editable draft and run a mandatory whole-figure convergence pass inside that same PowerPoint:

1. Inspect the entire page for local crowding, unused voids, asymmetric visual weight, and inconsistent scale.
2. Inspect every title band, text-to-border gap, text-to-line gap, image overlay, connector corridor, and neighboring module pair at 100–200%.
3. Resolve collisions by directly moving or resizing the existing objects, redistributing the complete peer group, expanding its content-driven box, or rerouting the existing path.
4. Recompute connectors after moving nodes; do not preserve a bad bend or split a route into independent line fragments merely to avoid re-layout.
5. Export a new preview from the same PPT and repeat the local adjustments until no accidental collision, border adhesion, cramped arrow label, avoidable bend, or visually sticky gap remains.

Prefer small coordinated movements that preserve alignment and logic. Do not independently nudge one peer, shrink one font, or shorten one arrow if the surrounding row or column can be rebalanced.

## Iterate normal figure tasks

Treat normal figure production as a two-phase workflow:

1. Generate one complete editable first draft containing all approved modules, text, formulas, images, symbols, and connections.
2. Keep that draft open as the authoritative artifact. Inspect the rendered result, then directly move, resize, crop, reorder, or reroute its existing objects in PowerPoint or through PowerPoint object control. Export another preview after each micro-adjustment pass.

Generate at most once for each confirmed structural design. Do not rerun the entire generator for overlap, whitespace, title-band, label-position, icon-scale, image-crop, arrow-routing, text-box-boundary, or z-order corrections. Perform all such corrections in the same PPT. Create a new structural version only when the user explicitly changes the master, system logic, module hierarchy, or content inventory; generate that new version once, freeze it, and resume direct object-level refinement.

Preserve stable object names and identities during refinement so collision reports and subsequent adjustments target the same objects. Treat the source script only as a one-time initial-construction tool for the current structural version, never as the normal mechanism for visual iteration.

Apply an “exactly one counted generation attempt per round, followed by read-only QA” restriction only when the user explicitly requests controlled batch-adaptation training and the frozen round contract requires it. In that mode, carry fixes into the next round and regenerate the complete test set. Never transfer that training restriction to an ordinary user figure task.

## Analyze the materials

Inventory the supplied files before reading deeply. Prefer `rg --files` and targeted searches for large folders. Read the relevant framework text, chapter outline, formulas, papers, captions, and reference images.

Before drawing, summarize the inferred design brief:

- Figure title.
- Main modules and submodules.
- Logical flow and hierarchy.
- Input/output relationships.
- Key technical points.
- Innovation highlights.
- Validation or experimental platform when relevant.
- Confirmed master, usable figure bounds, and intended figure aspect ratio.
- Typography tokens and content-editing boundary.
- Detected symmetry axes, repeated peer groups, and any justified asymmetry.

For a chapter-level overview, include motivation, chapter structure, research route, methods, key technologies, innovations, and validation platform when supported by the source materials. If the evidence does not support an item, do not invent it.

When multiple reference figures are supplied, fuse their strongest layout, hierarchy, and color ideas into one coherent design. Do not copy a single reference mechanically.

## Design the figure

Preserve all confirmed academic content. Improve layout rather than deleting modules or shrinking text aggressively.

Use these construction rules:

- Build the diagram with native PowerPoint shapes, text boxes, and connectors. Do not flatten the whole figure into an image inside the PPT.
- Start from the whole semantic object, not its visible fragments. Prefer one native PowerPoint rectangle, circle, diamond, star, line, elbow connector, curve, or polyline over a hand-built approximation.
- Before drawing any arrow, confirm the source and destination and therefore the arrowhead direction. Only then choose straight, elbow, curve, or diagonal geometry and set arrowhead type, length, width, and line weight. Do not add an arrowhead to an undirected source line.
- Classify every reference path before drawing it. Reproduce a source curve with a native PowerPoint curve or arc, a source folded path with one native elbow connector or polyline, and a source straight path with one straight line or connector. Never approximate a curve with straight segments or a folded path with several independent lines when one native object can express it.
- Prefer straight horizontal or vertical routing, then a balanced orthogonal fold, and only then a diagonal when the diagonal carries real meaning. Align anchors before selecting the route.
- Keep one continuous path as one object even when a factor, node, diamond, star, or other shape visually interrupts it. Draw the complete path on a lower z-order and place the intervening shapes above it. Occlusion is not a reason to split the path.
- Split a path only at a true branch, a junction with multiple semantic destinations, or independently attached endpoints. At a true branch, represent each unbranched run with the fewest native objects and place any junction marker above the lines.
- Establish z-order before construction: container backgrounds and frames at the back, continuous paths and connectors above them, factor/node shapes above paths, and labels plus annotations at the top. Verify that the overlay shape, not a white gap or manually shortened line, creates the visible interruption.
- Confirm whether every frame and peer shape uses a solid or dashed border before creating it. Apply the same border form to the matching main-diagram and legend representations.
- Treat repeated bars, dots, diamonds, and other matrix-like elements as one peer system. Lock one native shape type, size, interval, baseline, border form, and generation rule before placing the repeated set.
- Keep every block, label, formula, legend, boundary, and connector complete and within the canvas.
- Use native PowerPoint connectors with both endpoints genuinely attached when a straight or elbow connector can reproduce the intended route without uncontrolled rerouting. When a source curve or a fixed-waypoint folded route requires one native curve or freeform/polyline, place its endpoints exactly on the intended boundary anchors, name the path, and recompute its endpoints after either module moves.
- Prefer a single horizontal or vertical connector whenever endpoints can be aligned. Use orthogonal elbows for branches. Use diagonal connectors only when the diagonal direction carries real meaning.
- Align connector endpoints before choosing a connector type; do not introduce a bend to compensate for a few pixels of center misalignment.
- For a single folded route, use one native elbow connector or one native polyline and keep its bend location balanced. Use explicit waypoint anchors and multiple attached segments only for a true branch or when one native object cannot attach the required independent endpoints. Do not rely on uncontrolled automatic rerouting.
- Reserve connector and connector-label corridors before placement. Reject any connector segment that intersects text, a formula, a picture label, or another reserved object rectangle.
- Keep every route label in its own above-line or below-line corridor. Never center ordinary text directly on a connector, panel border, divider, or arrow shaft.
- Never attach a logical flow arrow to a nearby annotation text box when the intended endpoint is a module.
- Route connectors behind text and keep text above shapes when necessary.
- Use consistent line widths and medium arrowheads. Default to 1.0–1.2 pt unless the reference style requires otherwise.
- Preserve a clearly visible arrow shaft. Never let a medium arrowhead consume nearly the entire connector; size boxes, nodes, and gaps together and follow the routed-length checks in the quality checklist.
- Build decorative rules, ticks, dividers, and other nonlogical linework as thin rectangles or closed freeforms rather than `AddLine` objects that PowerPoint may reopen as connectors.
- Use distinct anchors for scientific vectors whose endpoints lie inside the same enclosing shape so PowerPoint does not collapse them into zero-length same-shape connectors.
- Build the layout from numeric guides: master margins, figure bounds, symmetry axes, panel rectangles, peer sizes, and repeated gaps. Do not position repeated or mirrored objects independently by eye.
- Use uniform rows, columns, module sizes, gaps, and page margins. Distribute whitespace evenly; avoid both crowded clusters and large unused voids.
- Normalize the apparent scale of peer icons, images, nodes, formulas, and labels before drawing connectors. Reject a peer group whose elements have inconsistent visual weight or whose local stacking creates a dense knot.
- Except for confirmed hard links that must touch their anchors, require visible whitespace between every symbol, label, frame, legend marker, annotation, and unrelated shape. Apply one gap token across peer items and measure edge-to-edge clearance, not center distance.
- Keep legend letters and labels visibly clear of their enclosing frame. Move the complete peer row or redistribute its contents when border clearance is tight; do not nudge one letter independently.
- For a small label placed in the channel between two large panels, calculate the channel first and give the label equal clearance to both panels when the structure is symmetric. If the clearance is too small but the master has spare room, move the paired panel groups apart symmetrically.
- Center sidebars against their associated panels. If a sidebar is taller than its panel, extend it equally beyond both ends; mirror opposite sidebars with the same rule.
- Keep at least 2.5–3 mm internal padding around text. Increase a box or adjust wrapping before reducing font size.
- Determine text-box width and height from measured text bounds before inserting text. Never fill text first and then stretch or distort the enclosing shape.
- Disable unexpected PowerPoint text-box auto-sizing before freezing coordinates. After the first render, compare the text box's actual `Shape.Height` with `TextRange.BoundHeight`; if PowerPoint retained an oversized default text-box height, set `TextFrame2.AutoSize` to none, shrink the shape to the rendered text height plus a safety allowance, and then reposition it. Do not assume the requested text-box height survived PowerPoint rendering.
- Treat a symbol and its adjacent explanation as a composite layout unit. Compute `symbol width + safety gap + measured label width`; start the label after the symbol's actual right or lower boundary, never from an approximate center or nominal coordinate. Use a visible 4–6 pt gap by default and apply the same gap across peer legend items.
- Before drawing a row or column, sum all occupied widths or heights, peer gaps, and parent padding. If the total does not fit, change the wording, module dimensions, or hierarchy before creating shapes.
- When unused space exists beside a crowded group, redistribute the whole peer group into that space before shrinking fonts, narrowing gaps, or allowing overlap. Preserve ordered alignment and equal peer gaps during the redistribution.
- Require every retained subordinate panel to contain at least one source-mapped mechanism, relation, or approved tightly scoped local image. An empty decorative frame is not a valid reconstruction.
- Keep block labels on one line whenever they fit with the required padding and a reasonable box width. Manual line breaks are permitted only for titles or content deliberately designed as exactly two lines; break at a semantic boundary and balance the two visual widths.
- For body text that naturally exceeds two lines, keep one continuous paragraph, allow wrapping from the fixed width, and apply full justification. Do not insert manual line breaks to imitate justification.
- Never leave one or two Chinese characters, a lone Latin variable, a subscript, an operator, or punctuation alone on a line. Fix the content width, wording, or box geometry while preserving the confirmed hierarchy.
- Center a title over its narrative column or left-align it to the same text edge; use one rule across the peer group. Apply the same paragraph alignment and usable width to peer descriptions so their left/right padding is equal.
- In addition to ordinary measurement clearance, require titles and peer labels to retain at least 10 pt spare inner width and 8 pt spare inner height in the target PowerPoint renderer. Before freezing geometry, serially probe the longest title and the widest and tallest peer labels with the final fonts and sizes.
- Optimize box widths and connector gaps jointly. Widening a box to fix wrapping must not reduce the adjacent connector to an arrowhead-only stub.
- Compute rounded-corner radii from the shorter side of each box and keep the normalized radius consistent within a peer group. Avoid pill-shaped corners unless the reference or hierarchy calls for them.
- Use only the user-confirmed number of font-size levels. Default to two restrained levels; use three or another count only when explicitly confirmed. Keep all sizes inside the confirmed range.
- When the active A4 batch contract specifies 12 pt and 14 pt, use exactly those two sizes. Do not silently generalize that batch-specific lock to a different confirmed contract.
- Apply the user-confirmed mixed-script font policy. If no policy is supplied after confirmation, default to Times New Roman for English, numbers, Latin variables, abbreviations, operators, and formulas, and 宋体 for Chinese.
- Keep formulas clear, ungarbled, and editable according to the confirmed formula contract. When MathType is selected, create genuine editable MathType OLE objects, preserve consistent baseline and apparent size, and verify the saved objects remain MathType-editable in PowerPoint. When Office-native equations are selected, use Office-native editable equations; otherwise use grouped editable text with superscripts, subscripts, fraction bars, brackets, and tables. Never substitute a formula screenshot unless the user explicitly authorizes a documented raster exception. Never silently mix formula technologies.
- Match the provided palette and visual language while maintaining academic restraint, adequate contrast, and print readability.

## Handle complex visual elements by scope

Classify every non-native visual before drawing and preserve its aspect ratio:

- Never paste the complete reference framework as a full-page or full-figure screenshot.
- For a large composition containing people, drones, sensors, computers, controllers, or other separable physical components, extract each component as an independent transparent object. Rebuild labels, dimensions, fields of view, module frames, and ordinary connectors natively in PowerPoint.
- Permit a tight crop for a small, subordinate, inseparable complex panel. Do not let that crop expand to include separable labels, boxes, legends, or major logical structure.
- Permit a lossless tight crop for a small blurry experimental plot, but do not enlarge it beyond a safe size, invent detail, or change its curves, axes, colors, or measurements.
- Faithfully crop and conservatively enhance real scenes, vehicles, people, hardware, RGB or thermal views, and experimental photographs. Do not redraw or semantically reinterpret them.
- Recreate small semantic icons as new transparent assets. Moderate detail differences are acceptable, but preserve their meaning, direction, dominant colors, silhouette, and visual weight.
- Render simple stars, circles, diamonds, squares, checks, crosses, and arrowheads as native PowerPoint shapes whenever semantics allow. Do not encode them as Unicode text merely for convenience.
- Generate irreducible local raster assets at no less than twice their final placed pixel dimensions, with transparency when appropriate. Reject blur, halos, loose crops, stretching, and source-semantic drift.

Use shape-center anchors as the first choice. Align center points before drawing arrows so paired connectors are exactly horizontal or vertical and symmetric spokes have equal routed lengths. Recompute the surrounding boxes rather than accepting a slightly tilted connector.

Reject these recurring failure modes:

- A node or oval and its explanatory text occupy intersecting rectangles even though empty space exists elsewhere in the parent panel.
- A label is positioned relative to a symbol's left coordinate or visual center and therefore intrudes into the symbol's actual width.
- Text passes the overflow check but overlaps another text box, frame, connector, node, or legend mark.
- One peer is moved independently to fix a collision, leaving unequal gaps or a broken symmetry pattern.
- Fonts are reduced or manual line breaks are added before testing whether available whitespace can absorb the content.
- Shapes are drawn first and content is forced into them without a measured occupancy preflight.

For a thesis figure, default to an A4-compatible canvas and size the final Word insertion to the usable text width. For journal or presentation use, adapt the aspect ratio only after confirming the target.

## Produce the three deliverables

Create exactly these final artifacts unless the user explicitly changes the contract:

1. Editable PowerPoint (`.pptx`) containing native, independently editable objects.
2. High-resolution PNG suitable for print, at least 300 dpi and preferably at least 2800 px on the long edge.
3. Word figure version (`.docx`) with the figure, caption, and legend placed together in a frameless, locked-anchor figure container when supported.

Keep temporary renders, scripts, and diagnostics in a separate build directory. Leave exactly the three deliverables in the final output directory. Ensure all three versions have the same visual result.

## Validate and refine

Read [references/quality-check.md](references/quality-check.md) before validating any generated figure. Perform both object-level checks and rendered-image inspection. If any check fails, revise the source PPT, export again, and repeat until the figure passes.

Do not finish after merely generating files. Do not report a connector as attached based only on visual proximity. Do not claim the PPT is editable if it contains only a full-slide raster image.
