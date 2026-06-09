# PowerPoint Tag Inspector

A Windows desktop application for viewing and editing PowerPoint Tags on the active slide or selected shape in a running PowerPoint presentation. Also allows renaming slides directly from the inspector.
<br/><img width="506" height="451" alt="image" src="https://github.com/user-attachments/assets/f5f78f15-eb09-454d-bf7f-f8f59cdc2700" />

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

1. Open Microsoft PowerPoint with a presentation.
2. Launch the PowerPoint Tag Inspector application.

### Main Window

The application window is divided into three sections:

#### 1. Options

Radio buttons to select the tag source:

- **Active Slide** — reads/writes tags on the currently active slide
- **Selected Shape** — reads/writes tags on the currently selected shape

The selected radio button displays contextual info (e.g., `Active Slide (#1 Slide1)` or `Selected Shape (#5 TextBox 3)`).

#### 2. Slide Name Editor

Located between Options and Tags. Shows:

- **Slide #N** label — the current slide index
- **Text field** — the editable slide name
- **Apply** button — commits the new name to PowerPoint

This works in both Active Slide and Selected Shape modes (always shows the current slide info).

#### 3. Tags

A table (ListView) displaying the tag list with columns **Name** and **Value**.

Action buttons:

| Button | Description |
|--------|-------------|
| **Add...** | Opens a dialog to create a new tag |
| **Edit...** | Opens a dialog to modify the selected tag's value |
| **Delete** | Removes the selected tag after confirmation |
| **Refresh** | Re-reads tags from the current target |

#### 4. Status Bar

Displays connection status, presentation name, mode, and tag count:

```
Connected: Demo.pptx | Mode: Active Slide | Tags: 3
```

### Tag Name Suggestions

When adding a new tag, the Name field offers autocomplete suggestions from a list of reserved tag names:

- `RowsDataSourceName`
- `RowsCount`
- `DeleteIf`
- `RowsGroupFieldName`
- `DeleteUnusedColumns`

You can still enter any custom name — the suggestions are non-restrictive.

### Error Handling

- If PowerPoint is not running, the app shows "Not connected" and allows retry via Refresh.
- If no shape is selected in "Selected Shape" mode, an appropriate message is displayed.
- All operations gracefully handle COM errors without crashing.

## Architecture

```
PowerPointTagInspector/
├── Program.cs                          # Application entry point
├── MainForm.cs / .Designer.cs          # Main UI (options, slide name editor, tags, status bar)
├── TagEditorDialog.cs / .Designer.cs   # Modal dialog for Add/Edit tag (with autocomplete)
├── Models/
│   ├── TagItem.cs                      # Immutable record: tag name + value
│   ├── TagSourceMode.cs                # Enum: ActiveSlide | SelectedShape
│   └── TagEditorState.cs              # State manager (mode, target, tags, slide info, errors)
└── Services/
	├── PowerPointService.cs            # COM Interop layer for PowerPoint
	└── ComInteropHelper.cs             # P/Invoke shim for GetActiveObject (.NET 8)
```

### Layers

| Layer | Responsibility |
|-------|----------------|
| **PowerPointService** | Direct COM interaction with PowerPoint (connect, get slides/shapes, read/write/delete tags, rename slides) |
| **TagEditorState** | Manages application state, coordinates between UI and service layer |
| **UI (MainForm + TagEditorDialog)** | User interface, event handling, displays errors |

### Key Technical Details

- **COM Interop in .NET 8**: `Marshal.GetActiveObject` was removed in .NET Core. A custom `ComInteropHelper` uses P/Invoke (`oleaut32.dll` / `ole32.dll`) to obtain the running PowerPoint instance.
- **PowerPoint Tags API**: Tags are accessed via `Slide.Tags` or `Shape.Tags` (1-based index). `Tags.Add(name, value)` performs upsert. Tag names are case-insensitive.
- **COM Object Lifecycle**: `PowerPointService` implements `IDisposable` and releases COM references on disconnect. The app never closes PowerPoint.

## Test Scenario

1. Open PowerPoint and create a new presentation.
2. Launch PowerPoint Tag Inspector.
3. Verify status bar shows: `Connected: Presentation1.pptx | Mode: Active Slide | Tags: 0`.
4. Verify the Slide Name editor shows `Slide #1` and the current slide name.
5. Change the slide name in the text field, click **Apply** — verify it updates in PowerPoint.
6. Click **Add...**, select `RowsDataSourceName` from suggestions, enter a value → OK.
7. Verify the tag appears in the list.
8. Select the tag, click **Edit...**, change the value → OK.
9. Verify the updated value is displayed.
10. Click **Delete**, confirm → tag is removed.
11. In PowerPoint, insert a shape and select it.
12. Switch to **Selected Shape** mode.
13. Add a tag to the shape, verify it persists.
14. Switch back to **Active Slide** — verify slide tags are shown (not shape tags).
15. Close PowerPoint → click **Refresh** → verify error message appears gracefully.
16. Reopen PowerPoint with a presentation → click **Refresh** → verify reconnection.

## Supported Versions

- Microsoft PowerPoint 2016, 2019, 2021, Microsoft 365
- .NET 8.0
- Windows 10 / 11
