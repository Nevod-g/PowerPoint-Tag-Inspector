using PowerPointTagInspector.Services;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointTagInspector.Models;

/// <summary>
/// Manages the current state of the tag editor: connection, mode, target object, and loaded tags.
/// </summary>
internal sealed class TagEditorState : IDisposable
{
    private readonly PowerPointService _service;
    private PowerPoint.Slide? _currentSlide;
    private PowerPoint.Shape? _currentShape;

    public TagEditorState()
    {
        _service = new PowerPointService();
    }

    /// <summary>
    /// Gets the current tag source mode.
    /// </summary>
    public TagSourceMode CurrentMode { get; private set; } = TagSourceMode.ActiveSlide;

    /// <summary>
    /// Gets whether the service is connected to PowerPoint.
    /// </summary>
    public bool IsConnected => _service.IsConnected;

    /// <summary>
    /// Gets the name of the active presentation.
    /// </summary>
    public string? PresentationName => _service.ActivePresentationName;

    /// <summary>
    /// Gets the currently loaded tags.
    /// </summary>
    public List<TagItem> CurrentTags { get; private set; } = [];

    /// <summary>
    /// Gets the last error message, or null if no error.
    /// </summary>
    public string? LastError { get; private set; }

    /// <summary>
    /// Connects to the running PowerPoint instance.
    /// </summary>
    public void Connect()
    {
        LastError = null;

        try
        {
            _service.Connect();
        }
        catch (InvalidOperationException ex)
        {
            LastError = ex.Message;
            throw;
        }
    }

    /// <summary>
    /// Sets the current mode and refreshes tags.
    /// </summary>
    public void SetMode(TagSourceMode mode)
    {
        CurrentMode = mode;
        Refresh();
    }

    /// <summary>
    /// Refreshes the current target and reloads tags based on the active mode.
    /// </summary>
    public void Refresh()
    {
        LastError = null;
        CurrentTags = [];
        _currentSlide = null;
        _currentShape = null;

        if (!IsConnected)
        {
            LastError = "Not connected to PowerPoint.";
            return;
        }

        try
        {
            switch (CurrentMode)
            {
                case TagSourceMode.ActiveSlide:
                    _currentSlide = _service.GetActiveSlide();
                    CurrentTags = _service.GetTags(_currentSlide);
                    break;

                case TagSourceMode.SelectedShape:
                    _currentShape = _service.GetSelectedShape();
                    CurrentTags = _service.GetTags(_currentShape);
                    break;
            }
        }
        catch (InvalidOperationException ex)
        {
            LastError = ex.Message;
        }
    }

    /// <summary>
    /// Adds or updates a tag on the current target.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when no valid target is available.</exception>
    public void SetTag(string name, string value)
    {
        EnsureTarget();

        switch (CurrentMode)
        {
            case TagSourceMode.ActiveSlide:
                _service.SetTag(_currentSlide!, name, value);
                break;

            case TagSourceMode.SelectedShape:
                _service.SetTag(_currentShape!, name, value);
                break;
        }
    }

    /// <summary>
    /// Removes a tag from the current target.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when no valid target is available.</exception>
    public void RemoveTag(string name)
    {
        EnsureTarget();

        switch (CurrentMode)
        {
            case TagSourceMode.ActiveSlide:
                _service.RemoveTag(_currentSlide!, name);
                break;

            case TagSourceMode.SelectedShape:
                _service.RemoveTag(_currentShape!, name);
                break;
        }
    }

    /// <summary>
    /// Checks whether a tag with the given name already exists on the current target.
    /// </summary>
    public bool TagExists(string name)
    {
        return CurrentMode switch
        {
            TagSourceMode.ActiveSlide when _currentSlide is not null => _service.TagExists(_currentSlide, name),
            TagSourceMode.SelectedShape when _currentShape is not null => _service.TagExists(_currentShape, name),
            _ => false
        };
    }

    /// <summary>
    /// Gets a formatted status string for the status bar.
    /// </summary>
    public string GetStatusText()
    {
        if (!IsConnected)
        {
            return "Not connected to PowerPoint";
        }

        if (LastError is not null)
        {
            return $"Error: {LastError}";
        }

        string presentationInfo = PresentationName ?? "Unknown";
        string modeText = CurrentMode == TagSourceMode.ActiveSlide ? "Active Slide" : "Selected Shape";

        return $"Connected: {presentationInfo} | Mode: {modeText} | Tags: {CurrentTags.Count}";
    }

    public void Dispose()
    {
        _service.Dispose();
    }

    private void EnsureTarget()
    {
        bool hasTarget = CurrentMode switch
        {
            TagSourceMode.ActiveSlide => _currentSlide is not null,
            TagSourceMode.SelectedShape => _currentShape is not null,
            _ => false
        };

        if (!hasTarget)
        {
            throw new InvalidOperationException("No valid target available. Refresh and try again.");
        }
    }
}
