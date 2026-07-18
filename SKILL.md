---
name: paper-fig-skill
description: Create publication-ready academic framework, technical-route, chapter-overview, and research-content figures as editable PowerPoint, with optional PNG and Word outputs. Use when Codex must reproduce or optimize a reference figure, extract research content, preserve editable text and equations, prepare complex local visuals, and build accurate semantic connectors. Use a 5–8 minute adaptive fast path for ordinary single-page reference reconstruction and a 15-minute multi-agent path for more complex figures.
---

# Paper-Fig-Skill

Build the figure once, merge parallel work once, and validate once. Preserve semantic correctness, editability, text clearance, formula editability, and connector direction while avoiding repeated analysis, repeated PowerPoint startup, and open-ended visual iteration.

## Use the 15-minute workflow by default

Target completion within 15 minutes after all source materials are locally available and the controlling contract is known. Treat this as the normal workflow, not a reduced-quality exception.

Preserve these non-negotiable requirements:

- No missing, clipped, unreadable, or accidentally overlapping content.
- Keep ordinary text and formulas clear of images, frames, modules, connectors, arrowheads, and unrelated text.
- Keep a source-horizontal or source-vertical route exactly horizontal or vertical.
- Attach connectors to the intended semantic object or explicit boundary anchor, never to a nearby label or decorative fragment.
- Keep required formulas editable under one confirmed formula technology.
- Keep logical structure, text, labels, frames, nodes, arrows, and connectors editable in PowerPoint.
- Never flatten the complete figure into a screenshot.

Use a high-assurance extended audit only when the user explicitly requests pixel-level reproduction, per-object equation reopening, or exhaustive verification. Do not let optional decorative differences trigger an extended audit.

## Use the 5–8 minute single-page fast path

Use this path when all of these are true:

- one reference image or one simple source slide;
- one output slide;
- no more than 60 visible text instances;
- no more than 80 logical routes;
- zero to three local raster crops;
- no newly generated scientific illustration;
- no more than four formulas, or no formulas;
- the user requests reference reproduction rather than a new research narrative.

Classify the task in no more than 20 seconds. Do not run every specialist merely because the responsibility exists:

- Skip the formula agent when the formula target list is empty.
- For one to four formulas, let the global agent insert them directly unless ribbon UI work would block layout.
- Merge text and asset work into one specialist when all raster work is limited to three lossless source crops.
- Start a dedicated complex-visual agent only for generated imagery, transparency work, segmentation, or more than three independent raster assets.
- Always assign arrow semantics separately when the figure contains more than 12 routes; otherwise the global agent may record the route ledger while laying out.

Keep the five logical responsibilities, but use only the agents that have nonempty work. Never spawn an agent to return an empty ledger.

Use this fast-path budget:

| Time | Work |
|---|---|
| 0:00–0:20 | Classify complexity, reuse the contract, and decide which roles are nonempty. |
| 0:20–2:00 | Run text/assets and arrow semantics in parallel while the global agent establishes the numeric skeleton. |
| 2:00–5:00 | Build the complete non-connector PPT once; do not export an intermediate preview. |
| 5:00–6:30 | In one PowerPoint session, apply native crops, mixed-script fonts, exact routes, arrowheads, z-order, and final text-box buffers. |
| 6:30–8:00 | Save once, render once, run `fast-ppt-audit.ps1` once, repair only blocking failures in the same session, and deliver. |

Use one PowerPoint process for the entire finalization pass. Do not open separate sessions to inspect pictures, apply crops, add routes, fix text, export the preview, and save. Collect picture indices, text metrics, route counts, and file validation before closing that one process.

For the base construction:

- Export the editable PPTX only. Do not render the artifact-tool first draft.
- Reserve 5% extra width for single-line labels and titles before construction.
- Create raster placeholders at their final rectangles, then apply native PowerPoint `PictureFormat.Crop` in the finalization session.
- Add all route records in one batch from explicit waypoints after geometry freezes.
- Apply fonts, crops, routes, and z-order in the same batch.

Render only the finalized slide. If that render exposes a blocking failure, modify the existing objects before closing PowerPoint and overwrite the same preview. Do not generate a new structural draft.

Run `scripts/fast-ppt-audit.ps1` as the default final check. Run `slides_test.py` or another full secondary audit only when the fast audit fails, cannot inspect the deck, or the user explicitly requests high assurance. Never run both by default.

On Windows, resolve the current user home once from the environment and keep `HOME`/`USERPROFILE` consistent before presentation-workspace setup. This prevents the artifact runtime from being rediscovered under the current drive. Never hard-code the resolved home in the skill or generated scripts.

## Run one compact intake and portable preflight

Reuse values already stated in the current request or conversation. Ask at most one compact confirmation only for missing, contradictory, or result-changing choices:

- master/canvas and orientation;
- Chinese and Latin fonts plus size tokens;
- formula technology and whether fallback is allowed;
- wording-preservation boundary and reference-fidelity level;
- requested outputs.

Do not ask the user to reconfirm an explicit value. Content extraction and visual classification may start while a nonblocking choice is unresolved; only geometry that depends on that choice must wait.

Resolve this skill from the directory containing `SKILL.md`. Never hard-code a username, drive letter, Codex home, repository path, MathType executable, or installation path. Treat `references/`, `scripts/`, and `agents/` as optional helpers.

Run capability detection once per task:

1. Record the available PowerPoint-writing surface.
2. Run `scripts/portable-preflight.py` when present; otherwise perform the equivalent minimal check.
3. When MathType is requested, confirm the visible PowerPoint MathType/WIRIS/Design Science ribbon add-in once and record whether `Equation.DSMT4` is registered.
4. Do not repeat capability, registry, add-in, or path discovery for each formula or each render.
5. Do not stop because an optional helper is absent. Use the self-contained workflow below.

## Coordinate five roles with one authoritative PPT

Use multi-agent work when at least two responsibilities are nonempty and can proceed independently. The main agent is the **global layout agent** and the only writer of the authoritative PPT. Other agents must not edit that PPT.

Treat these as responsibilities:

1. **Global layout agent:** own the canvas, coordinate system, panel/module geometry, object names, z-order, merge, final render, and acceptance decision.
2. **Text agent:** extract visible wording once and return final text records, typography tokens, wrapping rules, and measured occupied rectangles.
3. **Formula agent:** create the formula ledger and, when appropriate, a separate editable MathType formula-bank PPT.
4. **Complex-visual agent:** classify and prepare all crops, generated icons, transparent assets, and local raster exceptions in one batch.
5. **Arrow-logic agent:** map semantic routes in parallel, then resolve exact anchors and orthogonal waypoints after the non-connector layout freezes.

For a complex task with four concurrency slots, run the global layout agent plus the three largest nonempty specialist roles first. As soon as one specialist finishes, reuse that slot for the remaining nonempty role. Start arrow semantic planning early; delay physical connector creation until layout freeze. For a fast-path task, prefer two productive specialists over four low-work agents.

Use one-writer ownership:

- Specialists may read sources and create manifests, assets, or a test/formula-bank PPT.
- Specialists must not open or mutate the authoritative PPT.
- The global layout agent imports each specialist result once.
- Do not create several competing full-slide drafts.

Maintain one compact `figure-manifest.json` instead of three continuously updated prose ledgers. It must contain:

- canvas, margins, fonts, size tokens, gap token, and output contract;
- semantic object IDs, peer groups, rectangles, and z-layers;
- text records and formula targets;
- asset paths, crop rules, and target rectangles;
- route records, anchors, waypoints, and arrowhead ends.

Specialists may use separate JSON/TSV files during parallel work, but the global agent merges them once. Do not rewrite the manifests after harmless visual nudges; update only changed coordinates or content.

## Follow the fixed time budget

| Time | Work |
|---|---|
| 0:00–1:00 | Reuse/confirm the contract, run one capability preflight, establish canvas and stable object IDs. |
| 1:00–2:00 | Launch text, formula, and complex-visual work; global agent builds the panel/module skeleton; prepare arrow semantics. |
| 2:00–6:00 | Specialists produce text records, formula bank, assets, and logical route ledger while the global agent completes the non-connector layout. |
| 6:00–8:00 | Freeze occupied rectangles and anchors; resolve final formula boxes and route waypoints. |
| 8:00–11:00 | Merge text/assets/formulas once; add all connectors in one batch; apply z-order once. |
| 11:00–12:30 | Run one automated object audit and export one whole-slide preview. |
| 12:30–14:00 | Fix only reported blocking failures in the same PPT. Do not rebuild the structure. |
| 14:00–15:00 | Re-run failed checks, save/reopen once, and produce only requested deliverables. |

If one asset or formula is still unresolved at minute 6, continue the rest of the figure and repair only that item at merge time. Never restart completed parallel work because one item failed.

## Build the non-connector layout once

Inventory only the visible and logically necessary content. Avoid deep-reading unrelated files. Create stable IDs before construction and use numeric guides for margins, panels, rows, columns, symmetry axes, and peer gaps.

Place all panels, modules, text placeholders, formula placeholders, images, icons, labels, legends, and reserved connector corridors before drawing logical connectors. Normalize peer sizes and align the semantic whole object, not a visible fragment.

Use native PowerPoint shapes for logical nodes, frames, legends, ordinary symbols, labels, and editable text. Keep backgrounds at the back, route shafts above backgrounds, modules and nodes above routes, and labels/formulas at the top.

Use one default external `gap_token = 5 pt` unless the confirmed reference requires another value. Inflate each ordinary text/formula rectangle by that token for composition-level collision testing. Do not solve crowding by shrinking below the confirmed font size when usable space exists elsewhere.

Freeze the first complete non-connector layout. Ordinary spacing, z-order, crop, caption, or route corrections must edit the same PPT. Regenerate only when the user changes the canvas, hierarchy, semantic logic, or content inventory.

## Use the one-pass parallel text protocol

The text agent reads each source once and returns one record per visible text instance. Use this minimal schema:

```json
{
  "id": "TXT_MODULE_TITLE",
  "source_text": "Original wording",
  "display_text": "Original wording",
  "change_reason": null,
  "role": "module_title",
  "peer_group": "module_titles",
  "language": "en",
  "size_token": "title",
  "alignment": "center",
  "wrap_policy": "single_line",
  "allocated_width_pt": 110,
  "occupied_rect": [100, 40, 126, 30],
  "qa_flags": []
}
```

Preserve `source_text` and `display_text`; never silently shorten or paraphrase. Record an authorized change reason. Treat formula content as a formula placeholder and do not measure it as ordinary text.

Measure ordinary text once in one batch using final fonts, final sizes, the target PowerPoint renderer, and allocated widths. Disable unexpected AutoSize. Record rendered bounds, reject font substitution and overflow, and derive the text box with one consistent internal padding token. Permit a second measurement only if wording, font, size, renderer, or allocated width changes.

Use semantic manual breaks only for deliberately two-line titles or labels. Let body text wrap naturally. Flag orphan Chinese characters, isolated variables/operators, and punctuation. The global agent, not the text agent, runs the final inflated-rectangle collision test after all content is merged.

## Use the fast parallel MathType formula bank

When fewer than five formulas are required, insert them directly through the confirmed PowerPoint MathType ribbon workflow. When five or more are required, use a dedicated formula agent and a separate `mathtype_formula_bank_test.pptx`.

Create `formula-targets.json` once with ordered target names, TeX/content, slide numbers, reserved rectangles, and intended apparent size. The main agent places named placeholders and continues building without waiting.

Give the formula agent exclusive ownership of desktop PowerPoint UI during the batch. In one PowerPoint session:

1. Confirm the MathType ribbon once and insert one smoke-test formula.
2. Verify that the result is editable; when exposed, require `OLEFormat.ProgID = Equation.DSMT4`.
3. Insert all formulas into one labeled grid bank in ledger order.
4. Rename every equation to its exact target name.
5. Save at controlled checkpoints, not after every formula.
6. Mark an individual failed formula and continue; repair failures at the end.

After the bank and main layout are complete, open both in the same PowerPoint process and copy all formulas by name in one merge pass. Scale each equation proportionally into its reserved rectangle, center it, preserve its name, then remove the placeholder.

Run the default fast validation once: expected count, unique names, positive dimensions, slide bounds, editable OLE type/ProgID when exposed, and successful reopen of the bank and merged PPT. Reopen the first, middle, and last formulas through the MathType ribbon. Reopen every equation only when the user explicitly requests high assurance or a sample fails.

If the ribbon cannot be operated, use direct OLE only as a compatibility route when `Equation.DSMT4` is registered. Use Office-native equations or grouped editable math text only when the user authorizes fallback. Never use a formula screenshot without explicit permission.

Read `references/mathtype-tool.md` only when MathType is selected. Use `scripts/mathtype-ppt.ps1` when present for detection and saved-object validation.

## Prepare complex visuals in one batch

Classify each non-text visual in no more than 20 seconds:

- **Native:** use native PowerPoint when the object needs logical editing or can be expressed with at most three standard primitives.
- **Crop:** use one tight source crop for an inseparable subordinate photo, plot, scene, hardware view, or complex panel.
- **Generate:** generate only when no usable source exists, the asset carries necessary meaning, and native reconstruction would exceed 60 seconds.

Prepare all crops, resizes, background removals, and generated assets together while the main layout is built. Reuse an accepted asset; do not regenerate it for positioning changes. Cache derived assets in the build directory with stable names.

Require transparency only when the silhouette must overlap another object. Attempt background removal once. If cleanup would exceed 45 seconds or leaves a visible halo, use a clean rectangular crop inside a native frame.

A local raster is acceptable only when it is independently selectable and replaceable and contains no editable prose, formula, axis/tick label, logical connector, major module frame, or route label. Never rasterize the complete figure. Keep captions, labels, frames, and logical connections native.

If an asset is unusable by minute 6, replace it with a source crop in a native frame, a cached asset, or a neutral native symbol. Do not let one decorative asset block delivery.

## Plan arrow logic in parallel and draw once

The arrow-logic agent may analyze semantics while layout proceeds, but it must not draw provisional connectors. Return one route record per independently directed arrow:

```json
{
  "id": "R01",
  "source": "INPUT",
  "target": "MODULE_A",
  "source_anchor": "right",
  "target_anchor": "left",
  "arrow_end": "target",
  "kind": "straight",
  "waypoints": [[120, 80], [180, 80]],
  "segment_axes": ["H"],
  "line_pt": 1.1,
  "dash": "solid",
  "z_layer": "route",
  "min_span_pt": 12
}
```

Split every directed chain into separate routes when multiple arrowheads exist. Keep independent buses separate unless the source contains a true merge node. Record curves as curves and folded routes as one native elbow/polyline.

After the global agent freezes the non-connector layout, resolve named semantic anchors and exact waypoints once. Do not use a generic shape-center connector when the source specifies a boundary port or axis.

Before PowerPoint insertion, require:

- horizontal segment: `y1 == y2`;
- vertical segment: `x1 == x2`;
- orthogonal route: every segment is horizontal or vertical;
- endpoint lies on the declared semantic boundary anchor;
- arrowhead is on the ledger-declared destination end;
- shaft and arrowhead envelopes avoid inflated text/formula/image rectangles and unrelated modules;
- ordinary routed span is at least 12 pt;
- route IDs, source/target pairs, and visible arrow count are complete and nonduplicated.

Correct invalid JSON or frozen geometry before opening PowerPoint; never render a known diagonal or semantically misattached route for diagnosis. Add all route shafts and arrowheads in one batch and apply z-order once.

## Validate once and stop

Read `references/quality-check.md` before final validation when it exists. Use its fast gate, not an open-ended inspection loop.

When `scripts/fast-ppt-audit.ps1` exists, run it once against the authoritative PPT for slide bounds, text overflow, MathType count, and named `ROUTE_*` axis compliance. Treat its JSON as the initial failure list; do not repeat the same checks manually.

Run one automated object audit for:

- content and route count;
- slide bounds and nonempty objects;
- text overflow and font substitution;
- inflated text/formula clearance;
- route axis compliance, endpoint semantics, arrowhead end, and minimum shaft span;
- formula count, names, technology, dimensions, and bounds;
- raster aspect ratio and usable resolution;
- file save/reopen integrity.

Export one whole-slide preview. Inspect the whole composition and only connector-dense or audit-flagged regions at approximately 150%. Do not inspect every object repeatedly at 200–400%.

Fix blocking failures only: missing/clipped content, unreadable text, wrong flow direction, avoidable diagonal routes, text/object overlap, broken formulas, out-of-bounds objects, or unusable files. Apply one targeted same-PPT repair pass and rerun only failed checks plus one final render.

Stop when the automated gate has no blocking failures and the final render has no obvious collision or semantic error. Minor decorative, crop, icon-style, and pixel-level spacing differences do not trigger another iteration unless the user requested exact high assurance.

## Produce only requested deliverables

Do not create PNG or Word versions unless requested. If the user does not specify an output, deliver the editable PPTX and retain one QA preview only in the build directory.

Keep temporary manifests, formula banks, renders, scripts, and diagnostics outside the final delivery directory. Open or parse each requested final file once and confirm it is nonempty. Do not claim editability when the deck is a full-slide raster image.
