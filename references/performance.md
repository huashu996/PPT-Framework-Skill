# Fast, deterministic figure production

Use this playbook for complex reference reconstruction, MathType figures, dense connector networks, or any task where repeated rendering would dominate elapsed time.

## Default speed budget

Treat these as execution limits:

- one confirmation turn at most, and zero when the complete contract is already explicit;
- one merged analysis manifest rather than four repetitive narrative files;
- one structural generation and one authoritative PPT edited in place;
- one low-resolution structure render and one final PowerPoint render;
- one active PowerPoint session for structural post-processing and QA whenever possible;
- one sequential MathType editing session for all formulas;
- one combined MathType inspection and validation call;
- zero placeholder objects that will later be deleted and rebuilt with a different backend.

Exceed a limit only for a specific recorded failure. Do not render or regenerate merely to see whether an edit improved the result.

## Multi-agent collaboration

For a dense reference or an explicit multi-agent request, use the available concurrency once at the beginning:

1. Assign one agent exact text, symbols, and formula transcription with source rectangles.
2. Assign one agent the four-class inventory, independent-component boundaries, host-centered groups, and repeated-object counts.
3. Assign one agent the route ledger, arrow direction, path character, semantic anchors, shared baselines/axes, and likely collision corridors.
4. Keep the primary agent responsible for the canvas contract, manifest merge, construction backend, authoritative PPT, rendering, and final QA.

Every agent must return compact structured rows keyed by stable IDs. Do not ask agents to produce independent PowerPoints, duplicate the same inventory, or review one another sequentially. While agents run, the primary agent performs preflight, loads reusable helpers, and prepares the coordinate system. Join once, resolve conflicts in one merge pass, then author. If the figure is small or no concurrency is available, use the same schemas locally without delegation.

## Fast-path sequence

1. Resolve the canvas, typography, formula technology, content boundary, fidelity, and outputs once. Reuse the contract for all revisions of the same figure. If every value is explicit or the user says to start directly, do not stop for another confirmation.
2. Run environment preflight before full construction:
   - initialize the approved presentation runtime;
   - when MathType is selected, run `mathtype-ppt.ps1 -Action preflight`;
   - on a tiny scratch slide, verify the current wrapper's connector head/tail semantics, one horizontal connector, one vertical connector, one elbow connector, and rotated or vertically oriented text;
   - cache a successful result by host, PowerPoint version, MathType version, authoring-wrapper version, and skill version; rerun only when one key changes or the cached smoke test failed.
3. Analyze before authoring. Build one structured manifest containing the coverage, four-class, layout, and route fields. Store common facts once and expose four filtered views. Assign every non-text, non-formula element to exactly one construction class and one final backend. Give every independently movable raster component its own asset and placed image object; split multi-arrow or multi-icon artwork before construction. Speed is never a reason to omit an element, reduce a repeated count, flatten separable components, or change a source path's character.
4. Compile the numeric layout in the same manifest:
   - stable object names;
   - peer templates and repeated-shape generation rules;
   - shared horizontal baselines and vertical axes;
   - semantic connection sites;
   - fixed waypoints for routes that automatic PowerPoint rerouting may distort.
   - for every host-centered motif, place and measure the central text, image, formula, or node first, then derive every surrounding component from one recorded group center, visible-content radius, angular order, and clearance; never independently nudge orbiting components into approximate positions.
5. Choose the final authoring path before creating each object. If the final result requires PowerPoint Freeform/COM or OOXML crop/arrowhead handling, create it through that route in the first structural pass. Never create an approximate placeholder in one library and later delete and rebuild it with another.
6. Generate the editable structural draft once. Reuse component functions for repeated cubes, grids, cards, tables, activation stacks, legends, and peer modules. Batch all independent crops and all structural OOXML edits in one operation.
7. Keep one PowerPoint application and target presentation open. In that session, apply native freeforms/connectors, z-order, crop, route, and object-level fixes, then save.
8. Perform one low-resolution structure-only PowerPoint render before inserting MathType objects. Batch every discovered fix—text overflow, object bounds, peer alignment, arrow direction, axis locks, route crossings, z-order, proxy anchors, and source crops—then update the same PPT without another render when numeric checks are sufficient.
9. Finish all structural, crop, connector, freeform, and OOXML post-processing before formulas. Do not rewrite slide XML after MathType insertion unless repairing a confirmed defect.
10. Insert every MathType object, edit the named objects sequentially in one PowerPoint/MathType session, run `fit` once per object, save once, close once, and run `inspect-validate` once.
11. In one object-level QA pass, inspect connectors, axis-locked endpoint coordinates, arrowhead ownership, MathType ProgIDs, object bounds, text bounds, media counts, and output-directory contents.
12. Export the final PowerPoint render/high-resolution PNG once after object QA passes. Re-export only for a real defect visible in the final render.

## Backend and session rules

- Use the authoring library for objects it can create correctly in final form.
- Use PowerPoint COM/automation directly for native curves, freeforms, connection sites, or z-order behavior the library cannot preserve.
- Use OOXML only for deterministic batch edits that cannot be performed reliably through the current presentation API. Apply all such edits together before MathType insertion.
- Do not create a connector version of a curve and later replace it with a freeform. Create the final curve once.
- Do not depend on a preview renderer for source crops if it does not honor `srcRect`; use a single real PowerPoint structure render.
- Query a script's parameter schema or help once before its first call. Never discover a wrong parameter by running the expensive operation.
- Before ZIP/OOXML editing, verify that the target presentation is saved and not held open. If a PowerPoint session must remain active, perform OOXML edits first.
- Reuse the same active PowerPoint instance. Do not open a fresh application for insert, fit, inspect, validate, render, and audit separately.

## Host-centered batching

- Detect host-centered motifs during inventory, not during visual cleanup.
- Lock the host and one group-center coordinate first. Compute all surrounding positions in one batch from visible-content bounds, reference order, radius or edge clearance, and orientation.
- Keep each curved arrow or stylized component as its own asset/object, but store one group transform so the complete motif can be moved without repeated manual nudging.
- Place editable member labels in the same batch as their owning component.
- Render the completed motif once at structure resolution, correct its shared center/radius/order values, and avoid component-by-component render cycles.
- Do not open MathType while frames, host centers, member positions, or connectors are still changing.

## Connector lessons

- Treat `horizontal` and `vertical` as numeric constraints. One row uses one shared y baseline; one column uses one shared x axis.
- Compare the actual semantic anchor centers: use `Top + Height / 2` for horizontal side-center anchors and `Left + Width / 2` for vertical top/bottom-center anchors. Equal `Top` or `Left` values are not sufficient.
- Prefer real connection sites. If a proxy anchor is unavoidable, place it on the visible boundary and derive every peer proxy from the shared baseline, not from different bounding-box centers.
- Move modules or complete peer groups to align semantic anchors. Never preserve approximate module positions and accept a slightly diagonal route.
- Verify connector head/tail semantics with a scratch connector before full construction because wrapper naming may not match PowerPoint's visual begin/end terminology.
- Automatic elbows may change their last segment. When fixed waypoints or final-segment direction matter, use one native PowerPoint open polyline/freeform with its arrowhead on that same path.
- After PowerPoint routing, inspect actual endpoints and, for fixed paths, every waypoint segment. Horizontal segments must have equal endpoint y values; vertical segments must have equal endpoint x values.

## MathType lessons

- Preflight MathType and PowerPoint before drawing the full figure.
- Prefer the active PowerPoint instance when it is usable; never quit or close unrelated user presentations.
- Expect MathType to expand an edited OLE object to its intrinsic equation size. Use the bundled `fit` action to scale and center the same object inside its reserved rectangle.
- Enter formula content through the visible MathType editor, then save, fit, reopen, inspect, and validate. Do not substitute a screenshot or ordinary PowerPoint text.
- For unattended edits, focus the equation body, use `Ctrl+A`, press Backspace, then enter the replacement. Save with `Ctrl+S` and return with `Ctrl+F4`. The Backspace step prevents MathType from treating a selected equation as a formula-clipboard paste.
- Avoid multiline MathType content in a narrow vertical slot. Reserve a rectangle that matches the equation's natural aspect ratio. Keep nonformula variable inventories as adjacent native text only when the confirmed formula contract allows that separation.
- Edit every formula sequentially in one active PowerPoint/MathType session. Fit each object after closing its editor, but save and close the presentation only once after the complete batch.
- Run `mathtype-ppt.ps1 -Action inspect-validate -ExpectedCount N` once. Do not call separate `inspect` and `validate` processes.
- Render at structural and formula milestones, not after every formula. Never reopen MathType after final structural post-processing begins.

## Avoidable slow paths

- Positioning orbiting components individually before identifying the central host, then repeatedly correcting the apparent center and radial gaps.
- Using raw raster rectangles instead of alpha/content bounds and compensating for transparent padding through repeated manual nudges.
- Starting construction before the four-class inventory and reference coverage ledger are complete, then repeatedly discovering missing icons, arrows, or complex objects during QA.
- Reopening PowerPoint separately for each check instead of batching inspection in one session.
- Exporting a 300 dpi image after every small edit.
- Regenerating the complete presentation to correct spacing, labels, routing, or z-order.
- Discovering connector semantics, vertical text behavior, or MathType scaling only after the dense figure is built.
- Using approximate hidden anchors and then repairing many sloped connectors individually.
- Opening and fitting one MathType object in a fresh PowerPoint process for every formula instead of reusing one active presentation session.
- Re-rendering the full slide after each formula rather than after the completed formula batch.
- Writing four files that repeat the same object IDs, coordinates, classes, and status instead of one manifest with four views.
- Asking the user to reconfirm a complete contract or reconfirming unchanged values during revisions of the same figure.
- Drawing a temporary connector or raster approximation and later replacing it with the final PowerPoint curve, freeform, or crop.
- Running PowerPoint render, crop post-processing, connector audit, formula inspection, and validation in separate application sessions.
- Editing OOXML after MathType insertion and encountering a file lock or risking OLE relationship damage.
- Launching agents sequentially, giving them overlapping work, or letting several agents modify competing copies of the PPT.
- Calling a script before checking its parameter names, then repeating the operation after a preventable argument error.
