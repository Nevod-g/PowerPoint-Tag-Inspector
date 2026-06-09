# PowerPoint Tag Inspector — Design Document

A Windows desktop application that provides a simple UI for viewing and editing PowerPoint Tags on the active slide or selected shape in the currently active PowerPoint presentation.

## Technology Stack

| Component | Technology |
|-----------|-----------|
| Language | C# 12 |
| Framework | .NET 8.0 (Windows) |
| UI | Windows Forms |
| PowerPoint access | COM Interop via `Microsoft.Office.Interop.PowerPoint` |
| COM activation | P/Invoke (`oleaut32.dll` / `ole32.dll`) for `GetActiveObject` — required because `Marshal.GetActiveObject` was removed in .NET Core/.NET 5+ |

### Key Constraints

* The application works with an **already running** PowerPoint instance.
* It connects to the active presentation via COM `GetActiveObject("PowerPoint.Application")`.
* It never creates a new presentation without explicit user action.
* The primary scenario is reading and writing tags on the current active slide or the currently selected shape.

## Main UI

The application window contains two logical sections:

### 1. Options

Radio buttons for selecting the tag source:

* **Active Slide** — operates on `ActiveWindow.View.Slide`
* **Selected Shape** — operates on `ActiveWindow.Selection.ShapeRange[1]`

Switching the option immediately loads tags from the corresponding object.

### 2. Tags

A `ListView` (Details view) displaying the tag list with columns:

| Column | Description |
|--------|-------------|
| Name | Tag name (case-insensitive key) |
| Value | Tag value |

Available operations (buttons to the right of the list):

* **Add...** — opens a modal dialog to create a new tag
* **Edit...** — opens a modal dialog to modify the selected tag's value
* **Delete** — removes the selected tag after confirmation
* **Refresh** — re-reads the current target and reloads tags

### 3. Status Bar

A `StatusStrip` at the bottom shows:

```
Connected: Demo.pptx | Mode: Active Slide | Tags: 3
```

Or in error state:

```
Error: No shape selected
```

## Application Behavior

### Initialization

On startup the application:

1. Connects to the active PowerPoint instance via COM.
2. Obtains the active presentation.
3. Defaults to **Active Slide** mode.
4. Gets the active slide.
5. Loads and displays the slide's tags.

If PowerPoint is not running, no presentation is open, or the active slide is unavailable — a clear error message is shown in the status bar and action buttons are disabled.

### Active Slide Mode

1. Gets the active slide via `ActiveWindow.View.Slide`.
2. Reads all tags from `Slide.Tags` (1-based index).
3. Displays tags in the list.
4. Add/Edit/Delete operations target the active slide's tags.

### Selected Shape Mode

1. Gets the selection via `ActiveWindow.Selection`.
2. If `Selection.Type != ppSelectionShapes` — shows "No shape is currently selected in PowerPoint."
3. If multiple shapes are selected — uses the first one (`ShapeRange[1]`).
4. Reads all tags from `Shape.Tags`.
5. Displays tags in the list.
6. Add/Edit/Delete operations target the selected shape's tags.

## Tag Operations

### Add Tag

1. Opens `TagEditorDialog` with empty Name and Value fields.
2. Validates that Name is not empty (enforced on dialog close).
3. If a tag with that name already exists — asks for overwrite confirmation.
4. Calls `Tags.Add(name, value)` on the target (PowerPoint's `Add` performs upsert).
5. Refreshes the tag list.

### Edit Tag

1. User selects a row in the list.
2. Opens `TagEditorDialog` with current Name (read-only) and Value.
3. User modifies the value.
4. Calls `Tags.Add(name, newValue)` to update.
5. Refreshes the tag list.

### Delete Tag

1. User selects a row in the list.
2. Confirmation dialog is shown.
3. Calls `Tags.Delete(name)` on the target.
4. Refreshes the tag list.

### Refresh

Re-obtains the current PowerPoint object (slide or shape) according to the selected mode and re-reads all tags. This is critical because the user may switch slides or shapes directly in PowerPoint.

## Architecture

```
PowerPointTagInspector/
├── Program.cs                          — Application entry point
├── MainForm.cs / .Designer.cs          — Main window (options, tag list, buttons, status bar)
├── TagEditorDialog.cs / .Designer.cs   — Modal dialog for Add/Edit operations
├── Models/
│   ├── TagItem.cs                      — Immutable record: tag name + value
│   ├── TagSourceMode.cs                — Enum: ActiveSlide | SelectedShape
│   └── TagEditorState.cs               — State manager (mode, target, tags, errors)
└── Services/
	├── ComInteropHelper.cs             — P/Invoke shim for GetActiveObject (.NET 8+)
	└── PowerPointService.cs            — COM Interop layer for PowerPoint
```

### PowerPointService

Handles direct COM interaction.

| Method | Description |
|--------|-------------|
| `Connect()` | Obtains the running `PowerPoint.Application` via `GetActiveObject` |
| `Disconnect()` | Releases the COM reference |
| `GetActiveSlide()` | Returns `ActiveWindow.View.Slide` |
| `GetSelectedShape()` | Returns `Selection.ShapeRange[1]` |
| `GetTags(Slide)` / `GetTags(Shape)` | Reads all tags into `List<TagItem>` |
| `SetTag(target, name, value)` | Adds or updates a tag |
| `RemoveTag(target, name)` | Finds and deletes a tag by name |
| `TagExists(target, name)` | Checks if a tag exists (case-insensitive) |

### TagEditorState

Manages the current application state and coordinates between UI and service.

| Property / Method | Description |
|-------------------|-------------|
| `CurrentMode` | `ActiveSlide` or `SelectedShape` |
| `CurrentTags` | Currently loaded list of tags |
| `LastError` | Last error message (null if none) |
| `IsConnected` | Whether the service holds a COM connection |
| `Connect()` | Delegates to service, captures errors |
| `SetMode(mode)` | Switches mode and refreshes |
| `Refresh()` | Re-reads the target and its tags |
| `SetTag(name, value)` | Writes a tag to the current target |
| `RemoveTag(name)` | Deletes a tag from the current target |
| `TagExists(name)` | Checks for duplicate before add |
| `GetStatusText()` | Formats the status bar string |

### UI Layer (MainForm + TagEditorDialog)

Responsibilities:

* Mode selection (radio buttons)
* Tag list display (ListView)
* Add/Edit dialogs (TagEditorDialog)
* Error messages (MessageBox + status bar)
* Button enable/disable logic based on state

## Error Handling

All error scenarios are handled gracefully:

| Scenario | Behavior |
|----------|----------|
| PowerPoint not running | Status bar: "Cannot connect to PowerPoint…" |
| No active presentation | Status bar: "PowerPoint is running but no presentation is open." |
| No active slide | Status bar: "Cannot get the active slide…" |
| No shape selected | Status bar: "No shape is currently selected in PowerPoint." |
| Tag read failure | Error displayed in status bar |
| Tag write failure | MessageBox with details, then refresh |
| Tag delete failure | MessageBox with details, then refresh |
| User switches slide/shape in PowerPoint | Handled by Refresh button |

All COM exceptions are caught and wrapped in `InvalidOperationException` with user-friendly messages. The application never crashes due to COM errors.

## Technical Notes

### COM Interop in .NET 8

`Marshal.GetActiveObject` was removed in .NET Core. The application uses a custom `ComInteropHelper` class that calls `CLSIDFromProgID` (ole32.dll) and `GetActiveObject` (oleaut32.dll) via P/Invoke.

### PowerPoint Tags API

* Tags are accessed via `Slide.Tags` or `Shape.Tags` (1-based index).
* `Tags.Add(name, value)` performs upsert — if the name exists, the value is overwritten.
* `Tags.Delete(name)` removes a tag by name.
* `Tags.Name(index)` and `Tags.Value(index)` read individual entries.
* Tag names are **case-insensitive** in PowerPoint.

### COM Object Lifecycle

* The `PowerPointService` implements `IDisposable` and calls `Marshal.ReleaseComObject` on disconnect.
* The application does NOT quit or close PowerPoint — it only releases its COM reference.

## Acceptance Criteria

1. ✅ Detects a running PowerPoint instance with an open presentation.
2. ✅ Displays tags of the active slide.
3. ✅ Displays tags of the selected shape.
4. ✅ Can add a tag to the active slide.
5. ✅ Can edit a tag on the active slide.
6. ✅ Can delete a tag from the active slide.
7. ✅ Can add a tag to the selected shape.
8. ✅ Can edit a tag on the selected shape.
9. ✅ Can delete a tag from the selected shape.
10. ✅ Changes persist in the PowerPoint file (tags are written via COM, not cached locally).
11. ✅ Refresh correctly re-reads the current active slide or selected shape.
12. ✅ Graceful error messages when PowerPoint/presentation/shape is unavailable.

## Test Scenario

1. Open Microsoft PowerPoint and create or open a presentation.
2. Launch PowerPoint Tag Inspector.
3. Verify the status bar shows: `Connected: [filename] | Mode: Active Slide | Tags: 0`.
4. Click **Add...**, enter Name: `TestTag`, Value: `Hello` → OK.
5. Verify the tag appears in the list.
6. Select the tag, click **Edit...**, change value to `World` → OK.
7. Verify the updated value.
8. Click **Delete**, confirm → verify the tag is removed.
9. In PowerPoint, insert a shape and select it.
10. Switch to **Selected Shape** mode.
11. Add a tag to the shape, verify it appears.
12. Switch back to **Active Slide** — verify slide tags are shown (not shape tags).
13. Close PowerPoint, click **Refresh** — verify error message appears gracefully.
14. Reopen PowerPoint with a presentation, click **Refresh** — verify connection is restored.

## Supported Environments

* Windows 10 / 11
* .NET 8.0 SDK or Runtime
* Microsoft PowerPoint 2016, 2019, 2021, or Microsoft 365
