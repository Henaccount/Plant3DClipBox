# Plant3DClipBox

Sample Code, use at own risk: Concrete C# source bundle for a Plant 3D / AutoCAD 3D clip-box tool.

<table border=1><tr><td><i>Disclaimer – AI-Generated Code

This code was generated in whole or in part using artificial intelligence tools. While efforts have been made to review and validate the output, AI-generated code may contain errors, omissions, or unintended behavior.

Users of this code should be aware that:

The code may not follow best practices for security, performance, or reliability.
Vulnerabilities may be present, including but not limited to injection flaws, insecure dependencies, or improper handling of data.
The code may be susceptible to issues arising from prompt manipulation (e.g., prompt injection) or unintended generation of unsafe logic.
No guarantees are made regarding correctness, completeness, or fitness for a particular purpose.

It is strongly recommended that this code be thoroughly reviewed, tested, and audited—especially before use in production or security-sensitive environments.

The authors disclaim any liability for damages or issues arising from the use of this code.</i></td></tr></table>

<img src=P3DClipbox.png width=400>

## Included features

- Modeless palette UI based on `PaletteSet`
- Visible wireframe clip box on layer `P3D_CLIPBOX`
- `Fit selection` workflow
- Resize controls for X / Y / Z negative face, both faces, and positive face
- `All` resize control for symmetric growth or shrinkage in all directions
- `-` checkbox to force all delta fields negative for quick shrinking
- `move` checkbox to repurpose the six outer X / Y / Z buttons as box-translation controls along the clip-box local axes
- Highlighted move buttons while move mode is enabled, with the middle `<>` buttons and `All` resize action disabled in that mode
- Wider numeric input fields sized to show larger values more clearly, plus wider X/Y/Z/All labels and `<>` buttons in the palette
- Full summary and status text available on mouse hover via tooltips, with taller info panels and a wider default palette size
- Clipping on / off toggle
- Save / load / delete named clip-box states inside the DWG
- Native `MOVE`, `3DROTATE`, and `SCALE` workflow for the visible box when clipping is off

## Commands

- `P3DCLIPBOX` - show the palette

## Project files

- `Plant3DClipBox.csproj` for Plant 3D / AutoCAD 2025 and newer style builds
- `Plant3DClipBox.Net48.csproj` for Plant 3D / AutoCAD 2024 style builds

Update the `AutoCADSdkDir` property in the chosen project file if your installation path differs.

## Build and load

1. Open one of the project files in Visual Studio.
2. Confirm `AutoCADSdkDir` points to your Plant 3D installation folder.
3. Build the project.
4. In Plant 3D, run `NETLOAD` and load the built `Plant3DClipBox.dll`.
5. Run `P3DCLIPBOX`.

## Important behavior notes

### Fit selection from the palette

The palette button uses the current pickfirst selection. If nothing is preselected, use the command `P3DCLIPBOXFIT` from the command line and select objects there.

### Box visualization

The visible clip box is implemented as a reusable block reference of a unit wireframe cube. This keeps manipulation simple and robust.

### Position and orientation edits

Turn clipping off and then use normal AutoCAD commands such as `MOVE`, `3DROTATE`, and `SCALE` directly on the visible clip box. Press the palette `Refresh` button to re-read the box after manual edits.

### Move mode inside the palette

When `move` is checked, the six outer `<` and `>` buttons for X, Y, and Z move the clip box along its own local axes instead of resizing it. In move mode, the middle `<>` resize buttons and the `All` apply button are disabled, and the six move buttons are highlighted. The numeric fields are used as movement distances; the movement distance uses the absolute field value so the `-` checkbox remains a resize helper instead of flipping move direction.

### Minus checkbox

When `-` is checked, all delta fields are normalized to negative values. When it is unchecked, those same fields are normalized back to positive values. This makes it quick to switch between box growth and box shrinkage without retyping every field.

### Save / load storage

Named box states are stored in a shared user-profile XML store under the Autodesk application-data tree, so the same saved boxes are available from all drawings. For backward compatibility, older per-DWG states stored in the drawing Named Objects Dictionary under `P3D_CLIPBOX_STATES` are migrated into the shared store when encountered.

### Project reference check (`chk.ref`)

When a clip box is present, the `chk.ref` button checks whether the current drawing belongs to a Plant 3D project. If it does, the tool reads the Plant 3D drawing list from the active Plant project and scans those DWGs against the current clip box. For local projects, the drawing-path resolver now handles both absolute and project-relative Plant 3D folder paths, so drawings listed in Project Manager can still be found when the project uses relative paths. In Collaboration projects, the tool first checks whether all DWG files listed under Plant 3D Drawings in Project Manager are present in the local cache. If some are missing, you can cancel and download the missing files first, or continue with the cached files only.

The available candidate drawings are scanned against the current clip box through a hybrid path. If a project drawing is already open in Plant 3D, the tool scans that live database directly so an open drawing does not become a false candidate just because the file is locked for side-database reading. Otherwise the drawing is opened as a side database. A fast pass uses the DWG header model extents only as an early reject, then the drawing's own model-space entities are checked by extents. External-reference inserts inside those databases are excluded from that entity scan, and the tool's own clip-box helper block references on layer `P3D_CLIPBOX` are excluded as well, so a candidate drawing is judged by its own native model contents rather than by xrefs or leftover clip boxes it happens to contain. Conservative proposal is now limited to the case where no measurable local entity extents are available at all, which avoids false positives from stray unsupported helper entities when the measurable model contents are clearly far away.

The results window lists the missing Plant 3D project drawings that are likely relevant but are not already attached as xrefs. You can select all or individual rows and attach the chosen drawings as overlay xrefs at `0,0,0` using a relative path from the current host drawing. If clip-box clipping is already on, the tool reapplies clipping after loading so the newly added xrefs are clipped immediately.


### Clipping implementation

This bundle uses a hybrid clipping pass. Native top-level model-space entities are handled conservatively by hiding only objects whose extents are definitely outside the clip box, with a small safety margin so partially intersecting objects stay visible. Xref block references are clipped with an xclip-style spatial filter built from the clip box rectangle and its Z depth. The xref clip is recreated fresh on each apply, and the back depth is written in the negative extrusion direction so the front/back slice behaves like XCLIP clip depth. The clip box itself is hidden while clipping is on and shown again while clipping is off. Turning clipping off restores hidden native entities and restores or removes the temporary xref clips created by the tool.

## Files

- `src/ClipBoxPlugin.cs`
- `src/ClipBoxController.cs`
- `src/ClipBoxState.cs`
- `src/ClipBoxGsEngine.cs`
- `src/ClipBoxVisuals.cs`
- `src/ClipBoxStorage.cs`
- `src/ClipBoxControl.cs`
- `src/ClipBoxConstants.cs`
- `src/OrientedBox.cs`

## Limitations / follow-up ideas

- This source was authored against the documented managed AutoCAD API, but it was not runtime-tested here because Autodesk assemblies are not available in this environment.
- The current implementation is conservative by design: if an object extents intersects the box, the whole object remains visible. That avoids missing geometry inside the box.
- Native entities are still conservative: if an entity extents intersects the box, the whole entity stays visible. Xrefs now use a rectangular xclip-style volume based on the clip box and its depth.
- If an xref already has its own manual xclip, this tool temporarily replaces that xclip while clip-box clipping is on and restores the original xclip when clipping is turned off, before save, and on document close.
- If you later want exact face-level trimming, the next step would be deeper graphics-system clipping or geometry slicing.
