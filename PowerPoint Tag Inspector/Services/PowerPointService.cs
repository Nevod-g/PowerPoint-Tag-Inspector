using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointTagInspector.Services;

/// <summary>
/// Provides COM Interop access to a running PowerPoint instance.
/// Manages connection, tag reading, writing, and deletion.
/// </summary>
internal sealed class PowerPointService : IDisposable
{
    private PowerPoint.Application? _application;
    private bool _disposed;

    /// <summary>
    /// Gets the name of the active presentation, or null if not connected.
    /// </summary>
    public string? ActivePresentationName
    {
        get
        {
            try
            {
                return _application?.ActivePresentation?.Name;
            }
            catch (COMException)
            {
                return null;
            }
        }
    }

    /// <summary>
    /// Gets whether the service is connected to a running PowerPoint instance.
    /// </summary>
    public bool IsConnected => _application is not null;

    /// <summary>
    /// Connects to the active PowerPoint instance via COM.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when PowerPoint is not running or no presentation is open.</exception>
    public void Connect()
    {
        try
        {
            _application = (PowerPoint.Application)ComInteropHelper.GetActiveObject("PowerPoint.Application");
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException(
                "Cannot connect to PowerPoint. Ensure PowerPoint is running with an open presentation.", ex);
        }

        if (_application.Presentations.Count == 0)
        {
            throw new InvalidOperationException("PowerPoint is running but no presentation is open.");
        }
    }

    /// <summary>
    /// Disconnects from the PowerPoint instance without closing it.
    /// </summary>
    public void Disconnect()
    {
        if (_application is not null)
        {
            Marshal.ReleaseComObject(_application);
            _application = null;
        }
    }

    /// <summary>
    /// Gets the active slide from the active presentation.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when no active slide is available.</exception>
    public PowerPoint.Slide GetActiveSlide()
    {
        EnsureConnected();

        try
        {
            PowerPoint.DocumentWindow window = _application!.ActiveWindow;
            PowerPoint.Slide slide = (PowerPoint.Slide)window.View.Slide;

            return slide;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("Cannot get the active slide. Ensure a slide is selected in PowerPoint.", ex);
        }
    }

    /// <summary>
    /// Gets the first selected shape in the active presentation window.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when no shape is selected.</exception>
    public PowerPoint.Shape GetSelectedShape()
    {
        EnsureConnected();

        try
        {
            PowerPoint.Selection selection = _application!.ActiveWindow.Selection;

            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                throw new InvalidOperationException("No shape is currently selected in PowerPoint.");
            }

            PowerPoint.ShapeRange shapeRange = selection.ShapeRange;

            if (shapeRange.Count == 0)
            {
                throw new InvalidOperationException("No shape is currently selected in PowerPoint.");
            }

            return shapeRange[1];
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("Cannot get the selected shape. Select a shape in PowerPoint.", ex);
        }
    }

    /// <summary>
    /// Reads all tags from a slide.
    /// </summary>
    public List<TagItem> GetTags(PowerPoint.Slide slide)
    {
        ArgumentNullException.ThrowIfNull(slide);

        var tags = new List<TagItem>();
        PowerPoint.Tags slideTags = slide.Tags;

        for (int i = 1; i <= slideTags.Count; i++)
        {
            tags.Add(new TagItem(slideTags.Name(i), slideTags.Value(i)));
        }

        return tags;
    }

    /// <summary>
    /// Reads all tags from a shape.
    /// </summary>
    public List<TagItem> GetTags(PowerPoint.Shape shape)
    {
        ArgumentNullException.ThrowIfNull(shape);

        var tags = new List<TagItem>();
        PowerPoint.Tags shapeTags = shape.Tags;

        for (int i = 1; i <= shapeTags.Count; i++)
        {
            tags.Add(new TagItem(shapeTags.Name(i), shapeTags.Value(i)));
        }

        return tags;
    }

    /// <summary>
    /// Sets (adds or updates) a tag on a slide.
    /// </summary>
    public void SetTag(PowerPoint.Slide slide, string name, string value)
    {
        ArgumentNullException.ThrowIfNull(slide);
        ValidateTagName(name);

        slide.Tags.Add(name, value);
    }

    /// <summary>
    /// Sets (adds or updates) a tag on a shape.
    /// </summary>
    public void SetTag(PowerPoint.Shape shape, string name, string value)
    {
        ArgumentNullException.ThrowIfNull(shape);
        ValidateTagName(name);

        shape.Tags.Add(name, value);
    }

    /// <summary>
    /// Removes a tag from a slide by name.
    /// </summary>
    public void RemoveTag(PowerPoint.Slide slide, string name)
    {
        ArgumentNullException.ThrowIfNull(slide);
        ValidateTagName(name);

        PowerPoint.Tags tags = slide.Tags;
        int index = FindTagIndex(tags, name);

        if (index > 0)
        {
            tags.Delete(name);
        }
    }

    /// <summary>
    /// Removes a tag from a shape by name.
    /// </summary>
    public void RemoveTag(PowerPoint.Shape shape, string name)
    {
        ArgumentNullException.ThrowIfNull(shape);
        ValidateTagName(name);

        PowerPoint.Tags tags = shape.Tags;
        int index = FindTagIndex(tags, name);

        if (index > 0)
        {
            tags.Delete(name);
        }
    }

    /// <summary>
    /// Checks whether a tag with the given name exists on a slide.
    /// </summary>
    public bool TagExists(PowerPoint.Slide slide, string name)
    {
        ArgumentNullException.ThrowIfNull(slide);

        return FindTagIndex(slide.Tags, name) > 0;
    }

    /// <summary>
    /// Checks whether a tag with the given name exists on a shape.
    /// </summary>
    public bool TagExists(PowerPoint.Shape shape, string name)
    {
        ArgumentNullException.ThrowIfNull(shape);

        return FindTagIndex(shape.Tags, name) > 0;
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            Disconnect();
            _disposed = true;
        }
    }

    private void EnsureConnected()
    {
        if (_application is null)
        {
            throw new InvalidOperationException("Not connected to PowerPoint. Call Connect() first.");
        }
    }

    private static void ValidateTagName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Tag name cannot be empty.", nameof(name));
        }
    }

    private static int FindTagIndex(PowerPoint.Tags tags, string name)
    {
        for (int i = 1; i <= tags.Count; i++)
        {
            if (string.Equals(tags.Name(i), name, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return -1;
    }
}
