# Academic Figure Quality Check

Use this checklist after every generation or revision. A figure passes only when both the editable PPT and rendered PNG pass.

## 0. Hard priority gate

Reject the figure immediately if any of these conditions remains:

- An accidental overlap or occlusion exists between text, formulas, images, nodes, frames, arrows, labels, titles, legends, or panel boundaries.
- Any ordinary text's rendered rectangle, inflated by the confirmed external-clearance token or 4–6 pt by default, intersects an unrelated image, module, symbol, connector, arrowhead, divider, frame, boundary, or another text rectangle.
- An image caption touches or overlaps its image or enclosing boundary when unused space could separate the complete image-caption unit.
- A data route is diagonal even though the reference is horizontal/vertical or the endpoints can be aligned by moving the peer group or adding correct semantic anchors.
- A connector attaches to a decorative fragment such as one face of a 3D cube, one layer of a stacked image, or a label instead of the complete semantic object or its explicit boundary anchor.
- Two unrelated elements do not overlap numerically but are so close that they visually adhere.
- A title shares space with the first content row or a panel boundary.
- Ordinary text touches or crosses a connector, arrowhead, divider, frame, image edge, or module boundary.
- A connector label has no independent above-line or below-line corridor.
- Local crowding remains while usable space exists elsewhere in the same row, column, panel, or page.
- Peers use inconsistent sizes, gaps, alignment, or visual weight without a logical reason.

Apply the acceptance priority in this order: no accidental overlap, strong whitespace, balanced distribution, then positional similarity to the reference.

Do not count a confirmed source-semantic overlay as a collision when all of the following are true: the overlay is intentionally placed on an image or above a continuous lower-layer path; it remains inside the intended host; it has clear distance from the image and frame boundaries; and it does not cover unrelated content. Record every accepted overlay explicitly instead of globally disabling collision checks for its object type.

## 1. Content completeness

- Confirm the always-confirm intake gate recorded the master/canvas, typography contract, ordinary text size range, formula technology, formula size range, content-editing boundary, visual contract, and output contract for this invocation.
- Confirm the route ledger accounts for every source, destination, direction, junction, branch, merge, arrowhead end, and path type before accepting the visual reconstruction.
- Confirm the object/asset inventory classifies every whole visual object as a native PowerPoint object, editable equation object, tightly cropped source raster, or newly generated local raster.
- Compare the final labels against the approved design brief.
- Confirm every main module, submodule, formula, legend, boundary, arrow, input/output relation, key technique, innovation, and validation element that belongs in the figure is present.
- Confirm no research claim was added without support from the supplied materials.

## 2. Canvas, spacing, and alignment

- Confirm the master size and orientation were fixed before layout and match the user's explicit confirmation.
- Record the final figure bounding box, width/height ratio, usable margins, and orientation. Reject post-layout stretching or an upside-down/mirrored result.
- Keep every object inside the page boundary with consistent outer margins.
- Confirm a pre-layout occupancy record was made for every text, symbol, image, and connector label: locked font level, measured content bounds, padding, proposed rectangle, peer group, and parent panel.
- For every text object, record both the rendered glyph rectangle and its externally inflated clearance rectangle. Compare the inflated rectangle against every unrelated occupied object, including images, module borders, connector shafts, arrowheads, panel dividers, and enclosing frames; do not limit the check to neighboring text boxes.
- For every image caption, record the image bottom, caption top, caption rendered bottom, enclosing boundary, and both clear gaps. Reject the layout unless the caption is wholly below or beside the image as intended and clear of the enclosing boundary.
- Confirm the occupancy record was calculated jointly for text, formulas, symbols, icons, images, frames, arrows, arrowheads, and labels. Reject a layout assembled by placing one object category first and squeezing later categories into the leftover space.
- Before inspecting the rendered slide, verify that the sum of occupied peer widths or heights, repeated gaps, and parent padding fits the parent panel without intersection.
- Identify every left-right, top-bottom, radial, repeated-grid, or rotational structure. Record its axis/center and compare paired coordinates numerically, not only visually.
- For symmetric panel pairs, verify equal panel sizes, mirrored positions, equal outer margins, and equal central-channel clearance.
- For a small box between two large panels, measure the clear distance to both panel borders. Require equal clearance when symmetric and increase it by moving the paired panel groups apart if unused master space permits.
- Center each sidebar on its associated panel. If the sidebar is taller, verify equal extension beyond both panel ends and mirror the opposite sidebar with the same measurements.
- Check that same-row modules share a baseline and height; check that same-column modules share a center line or edge.
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
- When MathType is selected, insert and edit through the MathType plugin in the PowerPoint ribbon. After saving and reopening the PPT, select every equation and confirm that the ribbon edit command reopens it in MathType. Also run the dynamically resolved `scripts/mathtype-ppt.ps1 -Action inspect` and `-Action validate` commands when the helper is available. If it is missing, use the temporary build-directory adapter defined by the self-contained interface in `SKILL.md`; do not infer that installed MathType is unavailable. Confirm expected equation count, stable `MATH_*` names where supported, positive dimensions, and slide bounds. When an OLE ProgID is exposed, require `Equation.DSMT4`. When Office-native equations or grouped editable text are explicitly selected or authorized as fallback, verify that technology instead. Reject screenshots or silently mixed formula technologies unless the user explicitly approved a raster exception.

## 4. Connectors and arrows

- Confirm that a whole-object and connection ledger was completed before construction: semantic whole, native PowerPoint primitive, direction, path type, anchors, arrowhead ends, weight, dash style, and z-order.
- For every multipart object, confirm the ledger distinguishes the complete semantic object from its visible faces/layers and names the stable whole-object or explicit boundary anchors. Reject connectors attached to an arbitrary face, overlay, label, or subimage.
- Confirm the non-connector layout passed overlap, whitespace, boundary, peer-alignment, and reserved-corridor checks before logical connectors were added. Reject a workflow that drew arrows first and forced the modules around them.
- Verify arrow direction against the source or system logic before checking arrow type. Reject a correctly shaped arrow that points the wrong way, and reject an arrowhead added to an undirected source line.
- Confirm every source path was classified as straight, folded, curved, or truly branched before construction.
- Require one continuous unbranched path to be one native PowerPoint line, connector, curve, or polyline whenever the primitive can express it. Reject several independent straight segments that merely imitate one folded path or curve.
- Require source curves to be native PowerPoint curve/arc objects with continuous rendered curvature. Reject curves approximated by straight segments and reject straight or folded sources converted into curves.
- Require a single folded route to use one native elbow connector or polyline with balanced bend placement. Permit waypoint anchors and multiple segments only for a real branch or independently attached endpoints.
- When a node or factor interrupts a continuous path visually, verify that the complete path exists underneath and the intervening shape overlays it. Reject manually shortened or split line fragments used only to simulate occlusion.
- For true branches, verify that each unbranched run uses the fewest native objects and that any junction marker is layered above the branch lines.
- Use PowerPoint native connectors for straight and elbow routes whenever they preserve the intended geometry. Permit one native PowerPoint curve or fixed-waypoint freeform/polyline when the source requires a real curve or a stable folded route that automatic rerouting would distort.
- For connectors, verify `BeginConnected` and `EndConnected` are true. For an approved curve or freeform/polyline, verify both endpoint coordinates meet the intended module boundary anchors exactly and record that the path must be recomputed after either endpoint module moves.
- Inspect the names or text of connected shapes, or the recorded endpoint modules for approved native paths, to confirm every route reaches the intended module rather than a nearby label or invisible accidental object.
- Prefer a straight horizontal/vertical connector when both endpoints can be aligned.
- Compare the reference orientation before routing. If a source segment is horizontal/vertical or visually near that axis because of raster imprecision, require exact coordinate equality after reconstruction: equal endpoint y for horizontal routes and equal endpoint x for vertical routes. Any remaining avoidable diagonal is a layout failure.
- Compare shape-center anchors before routing. For symmetric pairs, require corresponding anchors to share the same x or y coordinate and require mirrored connectors to have equal routed lengths within normal rounding tolerance.
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
- Confirm repeated matrices use one consistent native primitive and numeric generation rule rather than individually improvised shapes.
- Confirm moving a module preserves its connector attachment; for approved native curve/freeform paths, confirm the same-file refinement pass recomputed their endpoints after the move.
- Reject a PPT that consists of a full-slide PNG or screenshot.
- Permit embedded raster source imagery only when explicitly required and when the diagram structure remains editable.

## 7. Rendered-image inspection

- Export the final slide/page to PNG at print resolution.
- Confirm there is exactly one authoritative editable draft for the current structural design. Modify that same PPT directly and preserve object identities for every visual correction; reject repeated generation.
- Inspect the entire figure at fit-to-page scale for balance and hierarchy.
- Overlay or mentally trace the recorded symmetry axes. Reject paired panels, sidebars, labels, central-channel boxes, or arrows that are visibly or numerically unbalanced.
- Inspect at 100% and approximately 200% for text clarity, arrow contact, clipping, and formula rendering.
- Inspect every connector-dense or annotation-dense region separately at 200–400%. At that scale, verify that ordinary data paths contain only the reference-approved orientations, text clearance remains visible on every side, captions remain detached from images, and arrow shafts remain visible.
- Run a numeric local audit after the visual pass: count non-orthogonal segments in routes classified as horizontal/vertical/orthogonal, report endpoint-coordinate deltas for straight routes, and report intersections between inflated text rectangles and all unrelated occupied objects. Require zero unexplained failures before delivery.
- Check all arrows for floating endpoints, oversized heads, arrowhead-only stubs, zero-length geometry, avoidable bends, and inconsistent weights.
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
