# Academic Figure Quality Check

Use this checklist after every generation or revision. A figure passes only when both the editable PPT and rendered PNG pass.

## 0. Hard priority gate

Reject the figure immediately if any of these conditions remains:

- An accidental overlap or occlusion exists between text, formulas, images, nodes, frames, arrows, labels, titles, legends, or panel boundaries.
- Two unrelated elements do not overlap numerically but are so close that they visually adhere.
- A title shares space with the first content row or a panel boundary.
- Ordinary text touches or crosses a connector, arrowhead, divider, frame, image edge, or module boundary.
- A normal logical arrow uses a separate arrowhead shape instead of one native PowerPoint line, connector, curve, or polyline carrying its own line-end arrowhead.
- A route classified as horizontal is not numerically horizontal, or a route classified as vertical is not numerically vertical.
- A connector label has no independent above-line or below-line corridor.
- A visible reference element is missing, merged into a simplified substitute, or represented by a generic placeholder without an explicit user-approved simplification.
- A non-text, non-formula element was drawn without a recorded choice among the four construction classes and a reason for that choice.
- One placed raster image contains several reference objects that could be moved, scaled, replaced, or layered independently, such as several arrows, icons, devices, nodes, or subsystem graphics.
- A host-centered structure was arranged before identifying and locking its central text/image/node, or its surrounding components have visibly unequal radius, clearance, ordering, or angular balance without support from the reference.
- A curved source object or arrow was changed to straight, a straight source object or arrow was changed to curved, an arrow direction changed, or repeated arrow/icon counts were reduced.
- Local crowding remains while usable space exists elsewhere in the same row, column, panel, or page.
- Peers use inconsistent sizes, gaps, alignment, or visual weight without a logical reason.

Apply the acceptance priority in this order: no accidental overlap, strong whitespace, balanced distribution, then positional similarity to the reference.

Do not count a confirmed source-semantic overlay as a collision when all of the following are true: the overlay is intentionally placed on an image or above a continuous lower-layer path; it remains inside the intended host; it has clear distance from the image and frame boundaries; and it does not cover unrelated content. Record every accepted overlay explicitly instead of globally disabling collision checks for its object type.

## 1. Content completeness

- Confirm the always-confirm intake gate recorded the master/canvas, typography contract, ordinary text size range, formula technology, formula size range, content-editing boundary, visual contract, and output contract for this invocation.
- Confirm the route ledger accounts for every source, destination, direction, junction, branch, merge, arrowhead end, path type, axis lock, and shared baseline/axis coordinate before accepting the visual reconstruction.
- Confirm every non-text, non-formula reference element is assigned exactly one of the four mandatory construction classes before drawing: original screenshot/cutout, self-generated imitation pasted into PowerPoint, one native PowerPoint shape, or stacked/grouped native PowerPoint shapes.
- Confirm the inventory records the selected class, reason, source rectangle, output object name, direction/orientation, curve/straight character, repetition count, z-order, and grouping boundary where applicable.
- For class 1, reject a whole complex screenshot when its visible members can be separated. For class 2, reject a generated composite placed intact when its members can be moved independently. For class 3, reject unnecessary decomposition of a PowerPoint-supported primitive. For class 4, verify the grouped native stack uses the fewest shapes that preserve the reference silhouette and z-order.
- Compare the reference coverage ledger against the final PPT one element at a time. Require every visible object and repeated icon to have a completed output counterpart; quality may vary, but unapproved omission is a hard failure.
- For generated or cropped raster layers, confirm separable text and formulas remain editable unless the user explicitly approved rasterizing them.
- For every class 1 or class 2 raster, confirm it contains exactly one independently movable semantic component unless the inventory documents a genuinely inseparable exception. For complex arrow motifs, require one placed image per curved or stylized arrow component and keep straight/simple arrows native when PowerPoint can reproduce them.
- For each host-centered structure, confirm the inventory names the host and records its visible bounds, visual center, surrounding component order, direction, ring/orbit radius, and clear gap.
- Verify the host is placed and measured before its surrounding components. Require the final host visual center to match the group center and require repeated/radial components to use consistent visible edge-to-host clearance unless the reference deliberately varies it.
- For cropped transparent components, inspect alpha/content bounds rather than the full image rectangle. Reject visually crooked placement caused by unequal transparent padding.
- Confirm every surrounding label is owned by and positioned relative to its host or member component, and that member-plus-label units preserve their spacing when the logical group moves.
- Confirm a later panel-level adjustment used one shared logical-group translation or recomputed transform. Reject independent member nudges that break the recorded center, radius, clearance, order, or symmetry.
- Compare the expected component ledger with the saved PPT: require one selectable placed image per independently movable raster member, require every embedded media file to be nonzero, and require stable names for the host and surrounding members.
- For stacked native objects, confirm the subshape count and z-order reproduce the reference construction and that the complete stack is grouped and named.
- Compare the final labels against the approved design brief.
- Confirm every main module, submodule, formula, legend, boundary, arrow, input/output relation, key technique, innovation, and validation element that belongs in the figure is present.
- Confirm no research claim was added without support from the supplied materials.

## 2. Canvas, spacing, and alignment

- Confirm the master size and orientation were fixed before layout and match the user's explicit confirmation.
- Record the final figure bounding box, width/height ratio, usable margins, and orientation. Reject post-layout stretching or an upside-down/mirrored result.
- Keep every object inside the page boundary with consistent outer margins.
- Confirm a pre-layout occupancy record was made for every text, symbol, image, and connector label: locked font level, measured content bounds, padding, proposed rectangle, peer group, and parent panel.
- Confirm the occupancy record was calculated jointly for text, formulas, symbols, icons, images, frames, arrows, arrowheads, and labels. Reject a layout assembled by placing one object category first and squeezing later categories into the leftover space.
- Before inspecting the rendered slide, verify that the sum of occupied peer widths or heights, repeated gaps, and parent padding fits the parent panel without intersection.
- Identify every left-right, top-bottom, radial, repeated-grid, or rotational structure. Record its axis/center and compare paired coordinates numerically, not only visually.
- For symmetric panel pairs, verify equal panel sizes, mirrored positions, equal outer margins, and equal central-channel clearance.
- For a small box between two large panels, measure the clear distance to both panel borders. Require equal clearance when symmetric and increase it by moving the paired panel groups apart if unused master space permits.
- Center each sidebar on its associated panel. If the sidebar is taller, verify equal extension beyond both panel ends and mirror the opposite sidebar with the same measurements.
- Check that same-row modules share a baseline and height; check that same-column modules share a center line or edge.
- Before adding connectors, require every horizontal route's semantic anchors to share one recorded y-coordinate and every vertical route's semantic anchors to share one recorded x-coordinate. If they do not, move the connected module or complete peer group and recompute the layout; do not compensate with a sloped connector.
- Compute those semantic anchors from actual centers. For horizontal side-center routes compare `Top + Height / 2`; for vertical top/bottom-center routes compare `Left + Width / 2`. Reject checks based only on equal `Top` or equal `Left` when peer sizes differ.
- Compare every peer group for identical width, height, normalized corner radius, padding, font level, alignment, wrapping rule, stroke, and repeated gap unless hierarchy requires a documented difference.
- Make repeated gaps equal unless hierarchy requires a deliberate difference.
- Compare whitespace above and below each content group. Redistribute modules when one large empty band remains.
- For every major panel, measure the title band separately from the content area. Require the title to clear the panel edge and the first content row, and require the first row's shapes, labels, and connectors to remain outside the title rectangle.
- Distinguish hard links from nonlinked neighbors. Require hard links to meet the intended anchors exactly and require all unrelated symbols, labels, annotations, and frames to retain a visible, repeated edge-to-edge gap.
- Measure label-to-symbol, legend-letter-to-frame, row-to-frame, and peer-to-peer clearances. Reject a row whose individual items fit but whose border clearance or gap rhythm is inconsistent.
- Keep at least 2.5–3 mm between text and its box border.
- Confirm final text and formulas were measured at their real font sizes before box geometry was derived. Reject fixed boxes that forced later font shrinking, crowding, or post-layout stretching.
- Compare every independent PowerPoint text box's actual shape height with its rendered `TextRange.BoundHeight`. Reject unexpected default-height or auto-size expansion that makes the text box collide with neighboring objects even when the glyphs appear smaller. Lock auto-sizing and shrink the text box to the rendered text plus safety padding before final collision checks.
- Prefer one-line block labels. Permit manual line breaks only for titles or deliberately two-line content, at semantic boundaries with balanced visual widths.
- For body text exceeding two lines, verify it is one continuous paragraph with automatic wrapping and full justification; reject manual line breaks used to simulate justification.
- Enumerate the rendered text lines for every wrapped label. Reject a line that contains only one or two Chinese characters, a lone Latin variable, a subscript, an operator, or punctuation; keep descriptors and variable lists together whenever practical.
- Verify peer narrative columns use the same usable width, left/right padding, heading alignment, and paragraph alignment.
- Adjust repeated box widths and connector gaps as one layout problem. A wider text box must not create an arrowhead-only connector in the neighboring gap.
- Keep clear separation between labels, formulas, arrows, frames, dots, grids, and illustrative geometry.
- Check every ordinary text label against line and border geometry, not only against other bounding boxes. Require the text to sit wholly above or below a connector or divider with visible clearance from the stroke and arrowhead.
- For every symbol-plus-label or node-plus-description unit, measure the gap from the symbol's actual boundary to the label's actual text box. Require a visible, consistent gap across peers; use 4–6 pt by default unless the confirmed reference requires another value.
- When a crowded group has unused space on the same row, column, or parent panel, reject font shrinking or overlap until the group has been redistributed into the available space with equal peer gaps.
- Do not fix crowding by deleting content or by sharply reducing font size.

## 3. Fonts and formulas

- Enumerate all font sizes in the PPT. Require exactly the user-confirmed number of size levels and keep every size inside the confirmed range.
- Confirm the formula technology matches the intake decision and that the formula size range matches the separately confirmed formula contract.
- Keep adjacent hierarchy levels restrained and consistent, normally 1–2 pt unless the user confirmed another scale.
- Inspect font runs, not only text-box defaults.
- Enforce the confirmed Chinese/Latin font policy. When the user accepted the default, apply Times New Roman to English, numbers, Latin variables, abbreviations, operators, and formula runs and 宋体 to Chinese.
- Reject Calibri, Aptos, or Cambria Math substitutions unless the user explicitly requests them.
- Check every formula for missing glyphs, empty objects, broken superscripts/subscripts, or encoding artifacts.
- When MathType is selected, run the bundled `scripts/mathtype-ppt.ps1 -Action inspect` and `-Action validate` commands after saving and reopening the PPT. Confirm every intended formula is a genuine `Equation.DSMT4` OLE object, the expected count and stable `MATH_*` names are present, and all formula objects remain inside the slide. When Office-native equations or grouped editable text are selected, verify that technology instead. Reject screenshots or silently mixed formula technologies unless the user explicitly approved a raster exception.
- Render every MathType object after `fit`. Reject empty objects, hairline-width objects, black-bar compression, malformed aspect ratios, formula-clipboard errors, or formulas that are technically present but unreadable. Re-enter the object with the delete-before-input sequence instead of covering the defect with ordinary text.
- Confirm that narrow variable inventories are not forcing a multiline MathType object into an incompatible portrait rectangle. If a neighboring native text inventory is used, confirm it is nonformula content allowed by the formula contract and that the actual mathematical expression remains a visible editable MathType object.

## 4. Connectors and arrows

- Confirm that a whole-object and connection ledger was completed before construction: semantic whole, native PowerPoint primitive, direction, path type, anchors, arrowhead ends, weight, dash style, and z-order.
- Confirm the non-connector layout passed overlap, whitespace, boundary, peer-alignment, and reserved-corridor checks before logical connectors were added. Reject a workflow that drew arrows first and forced the modules around them.
- Verify arrow direction against the source or system logic before checking arrow type. Reject a correctly shaped arrow that points the wrong way, and reject an arrowhead added to an undirected source line.
- Confirm every source path was classified as straight, folded, curved, or truly branched before construction.
- Confirm every straight route is also classified as `horizontal`, `vertical`, or meaningfully `free`. Horizontal and vertical classifications are hard constraints.
- Require one continuous unbranched path to be one native PowerPoint line, connector, curve, or polyline whenever the primitive can express it. Reject several independent straight segments that merely imitate one folded path or curve.
- Require every ordinary straight, elbow, curved, or fixed-waypoint arrow to be one native PowerPoint line/connector/path object whose arrowhead is set on that same object's beginning or ending line property. Reject a shaft plus a separate triangle, chevron, or arrowhead shape, including when the pieces are grouped.
- Require source curves to be native PowerPoint curve/arc objects with continuous rendered curvature. Reject curves approximated by straight segments and reject straight or folded sources converted into curves.
- Require a single folded route to use one native elbow connector or polyline with balanced bend placement. Permit waypoint anchors and multiple segments only for a real branch or independently attached endpoints.
- For an orthogonal elbow or fixed-waypoint polyline, require every intended horizontal segment to have equal endpoint y-coordinates and every intended vertical segment to have equal endpoint x-coordinates. Confirm the final segment approaches the destination from the direction recorded in the route ledger.
- When a node or factor interrupts a continuous path visually, verify that the complete path exists underneath and the intervening shape overlays it. Reject manually shortened or split line fragments used only to simulate occlusion.
- For true branches, verify that each unbranched run uses the fewest native objects and that any junction marker is layered above the branch lines.
- Use PowerPoint native connectors for straight and elbow routes whenever they preserve the intended geometry. Permit one native PowerPoint curve or fixed-waypoint freeform/polyline when the source requires a real curve or a stable folded route that automatic rerouting would distort.
- Do not accept an authoring-wrapper limitation as an exception. If the chosen library cannot assign an arrowhead to a PowerPoint-supported native path, switch to supported OOXML editing, PowerPoint automation, or another route that preserves the arrow as one native object.
- For connectors, verify `BeginConnected` and `EndConnected` are true. For an approved curve or freeform/polyline, verify both endpoint coordinates meet the intended module boundary anchors exactly and record that the path must be recomputed after either endpoint module moves.
- Inspect the names or text of connected shapes, or the recorded endpoint modules for approved native paths, to confirm every route reaches the intended module rather than a nearby label or invisible accidental object.
- Prefer a straight horizontal/vertical connector when both endpoints can be aligned.
- Prefer a straight horizontal/vertical connector whenever moving or aligning the modules can make the intended semantic anchors co-axial without changing the system logic.
- Compare shape-center anchors before routing. For symmetric pairs, require corresponding anchors to share the same x or y coordinate and require mirrored connectors to have equal routed lengths within normal rounding tolerance. When centers are misaligned, move the modules or complete peer group before drawing the connectors.
- Inspect the actual endpoint coordinates after PowerPoint has attached and rerouted each axis-locked connector. Require `|BeginY - EndY| <= 0.25 pt` for a horizontal route and `|BeginX - EndX| <= 0.25 pt` for a vertical route. Reject and realign the modules when either tolerance fails.
- For straight axis-locked connectors, also inspect the PowerPoint shape bounds: require `Shape.Height <= 0.25 pt` for horizontal routes and `Shape.Width <= 0.25 pt` for vertical routes. Run `scripts/ppt-axis-audit.ps1` when the connector names are known.
- For invisible proxy anchors, verify each proxy lies exactly on the intended visible boundary point. Require all proxy anchors in one horizontal peer row to use the same recorded y baseline and all proxy anchors in one vertical peer column to use the same recorded x axis. Reject anchors independently derived from different bounding-box centers.
- Before accepting an elbow, check whether moving an endpoint by a few pixels would allow a straight connector.
- Build branches from a clean orthogonal trunk and horizontal/vertical branches. Avoid diagonal fan-outs when a right-angle route is available.
- Use diagonal arrows only for meaningful geometric directions or scientific vectors.
- Keep line widths within 1.0–1.2 pt by default and use a consistent medium arrowhead.
- Compare arrowhead length and width as well as line weight. Reject a correct direction rendered with an oversized or inconsistent arrowhead.
- Measure the routed boundary-to-boundary connector length after `RerouteConnections`, not merely the pre-routing coordinates. With a medium arrowhead, target 14–18 pt of total connector span and reject spans below 12 pt for ordinary module-to-module arrows.
- In hub-and-spoke layouts, keep the radial gaps equal and large enough that every spoke has a visible shaft in addition to its arrowhead. Increase the panel, reduce the hub, or redistribute the surrounding boxes before shortening a spoke.
- Reject zero-length arrows and connectors whose beginning and ending shapes are identical unless an intentional self-loop is part of the logic. For scientific vectors drawn inside one enclosing shape, attach the endpoints to separate explicit anchors or scientific point objects.
- Put connectors behind labels and ensure no line crosses text or formulas.
- Confirm every connector label occupies a reserved above-line or below-line corridor. Reject labels centered on a shaft, placed against an arrowhead, or pressed against the source/destination boundary.

## 5. Overlap and clipping

- Detect objects outside the slide bounds.
- Check text overflow using the PowerPoint text bounds or equivalent object metrics.
- Inspect actual rendered text bounds and object bounding-box intersections, then visually exclude only intentional containment such as text inside its own shape. Do not exclude a text-versus-node intersection merely because both belong to one legend item.
- Check every text box against neighboring text, nodes, symbols, frames, images, and connector corridors. A clean overflow report is necessary but not sufficient.
- For peer legend items, record each symbol right/bottom boundary, label left/top boundary, and resulting clear gap. Reject negative or visually negligible gaps and compare peer gaps numerically.
- Reject any clipped box, text overflow, formula overlap, connector-through-text condition, or border collision.
- Confirm no object is hidden by an incorrect z-order.
- Inspect intentional overlays separately from accidental collisions. For every factor-over-line construction, require the path below, the factor/node above, and the label above both.
- For every symbol-over-image construction inherited from the source, require the symbol to remain inside the image's safe interior, clear of the image edge and enclosing frame, and above the raster in z-order.
- Confirm ordinary text boxes have no fill and no outline unless the reference requires a real labeled container.

## 6. PPT editability

- Confirm boxes, text, arrows, lines, and labels are independent native objects.
- Inventory native object types. Reject unnecessary decomposition when one PowerPoint basic shape, curve, elbow connector, or polyline can represent the semantic whole.
- Inspect each logical arrow's object identity. Require its shaft and arrowhead to belong to the same native line/connector/path object. Permit independent segmentation only after recording that PowerPoint itself has no native primitive for the required geometry, then use the fewest segments possible.
- Confirm repeated matrices use one consistent native primitive and numeric generation rule rather than individually improvised shapes.
- Confirm moving a module preserves its connector attachment; for approved native curve/freeform paths, confirm the same-file refinement pass recomputed their endpoints after the move.
- Reject a PPT that consists of a full-slide PNG or screenshot.
- Permit embedded raster source imagery only when explicitly required and when the diagram structure remains editable.

## 7. Rendered-image inspection

- Export the final slide/page to PNG at print resolution.
- Confirm there is exactly one authoritative editable draft for the current structural design. Modify that same PPT directly and preserve object identities for every visual correction; reject repeated generation.
- Inspect the entire figure at fit-to-page scale for balance and hierarchy.
- Overlay or mentally trace the recorded symmetry axes. Reject paired panels, sidebars, labels, central-channel boxes, or arrows that are visibly or numerically unbalanced.
- For every host-centered motif, overlay the recorded group center and inspect the host, surrounding visible-content centers, member labels, angular order, and outer visible bounds as one composition.
- Inspect at 100% and approximately 200% for text clarity, arrow contact, clipping, and formula rendering.
- Check all arrows for floating endpoints, oversized heads, arrowhead-only stubs, zero-length geometry, avoidable bends, and inconsistent weights.
- At 100–200% local view, inspect every axis-locked route for residual slope, then compare its numeric endpoint coordinates. Do not accept a line that merely appears almost horizontal or almost vertical at fit-to-page scale.
- Check every wrapped label for semantic balance and orphaned single-character or single-variable lines.
- Check all boxes for internal padding and all sections for even whitespace. When the master has unused vertical or horizontal space, reject avoidably tight channels between labels and large panels.
- Perform repeated whole-page and local micro-adjustment passes in the same PPT after the detailed checks. Move and resize existing objects, rebalance crowded groups, unequal gaps, scale mismatches, and title/content spacing, reroute affected connectors, then export and inspect again.
- Permit one new generation only when the user explicitly changes the master, system logic, hierarchy, or content inventory. Record the structural reason, freeze that new draft immediately, and return to direct PowerPoint refinement.
- Compare the PNG against the PPT and Word version for font or layout drift.
- If any issue appears, modify the editable source, re-export, and repeat the checks.

## 8. Final output contract

- Keep only three files in the final delivery directory:
  1. Editable PPT.
  2. High-resolution PNG.
  3. Word figure version.
- Do not include previews, scripts, PDFs, logs, or intermediate renders in the final directory.
- Open or parse each final file once to ensure it is readable and non-empty.
- Close the target presentation before cleanup so PowerPoint lock files are removed. Verify every recursive cleanup target resolves inside the intended output directory, then leave only the artifacts requested by the current output contract.
