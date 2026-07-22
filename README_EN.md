# PPT Framework Skill

English | [中文](README.md)

A Codex skill for creating editable academic framework diagrams in PowerPoint. It can faithfully reconstruct a reference image as an editable PPTX, or design a new technical roadmap, research framework, system architecture, or chapter overview from a paper, Word document, outline, or structured brief.

The goal is not to produce a similar-looking bitmap. The deliverable is a PowerPoint file built from editable native objects: text boxes, shapes, connectors, arrows, images, and editable Office-native or real MathType equations.

## Core capabilities

- **Faithful reference reconstruction**: Preserve the visible modules, text, shapes, curves, arrows, colors, layers, and connection topology without approximate redesign or simplification.
- **Design from source content**: Read a paper, Word document, outline, or written structure; create a low-fidelity SVG skeleton for approval before building the final PPT.
- **Native editability**: Do not use a full-slide screenshot as a fake editable result. Text, nodes, logical arrows, and equations remain independently editable.
- **Strict arrow geometry**: Preserve arrow type, head count, direction, anchors, and route. Every elbow arrow must remain one native PowerPoint object—never a stitched set of line segments.
- **Text layout constraints**: Keep body text as real paragraphs; prevent arbitrary hard breaks, one-character lines, isolated English tokens, and text collisions.
- **Two equation modes**: Delegate to the latest `formula-skill`, supporting PowerPoint native professional equations and genuine MathType `Equation.DSMT4` objects without substituting one for the other.
- **Mandatory equation confirmation**: Before recognition or insertion, confirm the equation type, full target PPTX path, and one-page or multi-page layout.
- **Fast single-slide workflow**: A normal one-slide build targets roughly ten minutes using one build, one local repair, and one final inspection.

## Demos

### 1. Faithful reconstruction of a dense technical roadmap

The original hierarchy, palette, local images, and layout are preserved while the main framework elements are rebuilt as editable PowerPoint objects.

![Faithful reconstruction of a dense technical roadmap](docs/images/demo-faithful-redraw.png)

### 2. Neural-network architecture with a MathType equation

The multi-scale spatio-temporal architecture keeps strict horizontal and vertical connections, module boundaries, and a real editable MathType equation.

![Neural-network architecture with MathType](docs/images/demo-mathtype-architecture.png)

### 3. Framework redesigned from research content

Research background, problem statement, framework, technical route, and study modules are reorganized into a clear, collision-free academic diagram.

![Framework redesigned from research content](docs/images/demo-from-outline.png)

### 4. Complex cyclic and multi-region framework reconstruction

The workflow handles cyclic mechanisms, curved arrows, four experiment regions, vertical side labels, and a dense central model.

![Complex cyclic and multi-region framework reconstruction](docs/images/demo-complex-cycle-redraw.png)

## Two operating modes

### Mode A: faithful reference reconstruction

Use this mode for requests such as:

- “Recreate this figure as an editable PPT.”
- “Redraw this paper figure in PowerPoint.”
- “Keep the original aspect ratio, typography, and arrow topology.”

Workflow:

1. Summarize page size, orientation, aspect ratio, and font range; when equations are present, also confirm equation type, full PPTX path, and pagination.
2. Wait for final confirmation of all production parameters.
3. Record every source object, label, arrow, anchor, layer, and geometric relationship.
4. Generate one authoritative PPTX.
5. Use the latest `formula-skill` to batch-insert the confirmed editable equation type when needed.
6. Render the final saved PPTX, repair blocking issues, and deliver only after validation.

The reference image already defines the structure, palette, and content, so the skill does not redundantly ask for a template, color scheme, or content source.

### Mode B: design from a paper, outline, or written structure

Use this mode for:

- Technical roadmaps
- Research frameworks
- System architecture diagrams
- Chapter arrangement diagrams
- Paper methodology overviews

This mode has two independent approval gates:

1. **Parameter approval**: Confirm the template or structural style, palette, source content, content-processing policy, page, typography, and equation plan.
2. **SVG skeleton approval**: Present a “PPT skeleton preview (not the final PPT)” showing regions, modules, flow, feedback, and equation positions.
3. Open PowerPoint and build the editable PPTX only after the complete skeleton is approved.

## Highest-priority rules

1. Visible glyphs must not collide with borders, arrows, shapes, images, equations, or other text.
2. Every elbow line or elbow arrow must be one editable native PowerPoint object; stitched segments are forbidden.
3. Unless explicitly present in the source, horizontal text must not contain one-character lines, isolated tokens, punctuation-only lines, or bullet-only lines.
4. Reference reconstruction must not remove, merge, rewrite, or invent modules, text, shapes, curves, or arrows.
5. Source connections that are horizontal or vertical must remain strictly horizontal or vertical, with the original anchor logic preserved.
6. When a PPTX path is known, PowerPoint must be launched directly with that absolute path. Do not use screenshots to find an Open dialog and type the path.
7. Delegate all equation recognition and insertion to the latest `formula-skill`; Office mode must produce native equations and MathType mode must produce `Equation.DSMT4`.

See [SKILL.md](SKILL.md) for the complete controlling workflow.

## Installation

### Option 1: clone into the Codex skills directory

```powershell
git clone https://github.com/huashu996/PPT-Framework-Skill.git `
  "$HOME\.codex\skills\Paper-framework-skill"
```

Restart Codex and invoke:

```text
$paper-fig-skill
```

### Option 2: update an existing installation

```powershell
git -C "$HOME\.codex\skills\Paper-framework-skill" pull
```

## Usage examples

### Faithfully reconstruct a reference image

```text
$paper-fig-skill Recreate this image as an editable PPT.
A4 portrait, preserve the original aspect ratio, SimSun 10–14 pt,
and the source contains no equations.
```

### Reconstruct an academic figure with equations

```text
$paper-fig-skill Recreate this figure in PowerPoint.
A4 landscape, Times New Roman 10–16 pt, use MathType for equations,
save to C:\output\framework-mathtype.pptx, and keep everything on one slide.
```

### Design a framework from a thesis chapter

```text
$paper-fig-skill Create a system framework diagram from Chapter 3
of the provided Word document. Use the supplied reference image as
the structural style, A4 portrait, a blue-green palette, Office-native
equations, C:\output\framework-office.pptx, and a one-slide layout.
```

The skill first summarizes all production parameters. For design-from-content tasks, it also presents an SVG skeleton and waits for approval before creating the formal PowerPoint.

## Requirements

- Codex with access to this project's `SKILL.md`
- Microsoft PowerPoint desktop on Windows
- Python 3 and `pywin32` (`python -m pip install -r requirements.txt`)
- For equations: the latest `formula-skill` installed locally
- MathType installed and registered as `Equation.DSMT4` only when MathType mode is selected; Office-native mode does not require MathType
- The Codex presentation runtime and `@oai/artifact-tool`
- For reference reconstruction: a sufficiently clear PNG or JPG source

MathType is optional and is used only when the user explicitly selects MathType. Equation recognition and insertion follow the sibling installation's `formula-skill/SKILL.md`.

## Repository structure

```text
PPT-Framework-Skill/
├── SKILL.md
├── README.md
├── README_EN.md
├── requirements.txt
├── agents/
│   └── openai.yaml
└── docs/
│   └── images/
│       ├── demo-faithful-redraw.png
│       ├── demo-mathtype-architecture.png
│       ├── demo-from-outline.png
│       └── demo-complex-cycle-redraw.png
```

Equation handling is no longer duplicated in this repository; it is delegated to the standalone latest `formula-skill`.

## Design principles

- Speed comes from parameterized layout, reusable components, and reduced rework—not from lower fidelity.
- Explicit user approval overrides generator defaults.
- Visible source geometry overrides semantic guesses and automatic connector routing.
- The final saved PowerPoint state overrides generator previews, window thumbnails, or object counts.
- The skill does not claim “reviewed” or “deliverable” until the final checks pass.

## Scope

This skill is optimized for dense academic framework diagrams on one or a small number of slides. It is not a general long-form presentation generator. Experimental plots, microscopy images, and photographs may remain as independently embedded images, while the logical framework should remain editable.
