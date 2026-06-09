# PowerPoint Tag Inspector

A Windows desktop application for viewing and editing PowerPoint Tags on the active slide or selected shape in a running PowerPoint instance.

## Requirements

- Windows 10 or later
- .NET 8.0 SDK or Runtime
- Microsoft PowerPoint (2016 or later) installed and running

## Building

```bash
dotnet build
```

## Running

```bash
dotnet run --project "PowerPoint Tag Inspector"
```

Or open the solution in Visual Studio and press F5.

## Usage

### Prerequisites

1. Open Microsoft PowerPoint with a presentation
2. Launch the PowerPoint Tag Inspector application

### Workflow

1. **On startup**, the application connects to the running PowerPoint instance and loads tags from the active slide.
2. Use the **radio buttons** to switch between:
   - **Active Slide** — reads/writes tags on the currently active slide
   - **Selected Shape** — reads/writes tags on the currently selected shape
3. Use the **action buttons**:
   - **Add...** — opens a dialog to add a new tag (name + value)
   - **Edit...** — opens a dialog to modify the selected tag's value
   - **Delete** — removes the selected tag after confirmation
   - **Refresh** — re-reads tags from the current target (useful after switching slides/shapes in PowerPoint)
4. The **status bar** shows connection status, presentation name, mode, and tag count.

### Error Handling

- If PowerPoint is not running, the app shows "Not connected" and allows retry via Refresh.
- If no shape is selected in "Selected Shape" mode, an appropriate message is displayed.
- All operations gracefully handle COM errors without crashing.

## Architecture

```
PowerPointTagInspector/
├── Program.cs                          # Application entry point
├── MainForm.cs / .Designer.cs          # Main UI (radio buttons, ListView, buttons, status bar)
├── TagEditorDialog.cs / .Designer.cs   # Modal dialog for Add/Edit tag
├── Models/
│   ├── TagItem.cs                      # Immutable record for tag name+value
│   ├── TagSourceMode.cs                # Enum: ActiveSlide | SelectedShape
│   └── TagEditorState.cs               # State manager (mode, target, tags, errors)
└── Services/
	├── PowerPointService.cs            # COM Interop layer for PowerPoint
	└── ComInteropHelper.cs             # P/Invoke shim for GetActiveObject (.NET 8)
```

### Layers

| Layer | Responsibility |
|-------|----------------|
| **PowerPointService** | Direct COM interaction with PowerPoint (connect, get slides/shapes, read/write/delete tags) |
| **TagEditorState** | Manages application state, coordinates between UI and service layer |
| **UI (MainForm + TagEditorDialog)** | User interface, event handling, displays errors |

## Test Scenario

1. Open PowerPoint and create a new presentation
2. Launch PowerPoint Tag Inspector
3. Verify status bar shows: `Connected: Presentation1.pptx | Mode: Active Slide | Tags: 0`
4. Click **Add...**, enter Name: `Author`, Value: `John` → OK
5. Verify the tag appears in the list
6. Select the tag, click **Edit...**, change Value to `Jane` → OK
7. Verify the updated value is displayed
8. Click **Delete**, confirm → tag is removed
9. In PowerPoint, select a shape on the slide
10. Switch to **Selected Shape** mode
11. Add a tag to the shape, verify it persists
12. Switch back to **Active Slide** — shape tags are replaced by slide tags
13. Close PowerPoint → click Refresh → verify error message appears gracefully

## Supported Versions

- PowerPoint 2016, 2019, 2021, Microsoft 365
- .NET 8.0
- Windows 10/11
