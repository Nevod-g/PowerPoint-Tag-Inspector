using System.Runtime.InteropServices;
using Core = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointTagInspector.Services;

/// <summary>
/// Provides COM Interop access to a running PowerPoint instance.
/// Manages connection, tag reading, writing, and deletion.
/// </summary>
internal sealed class PowerPointService : IDisposable
{
    private const int RpcCallRejectedMaxRetries = 5;
    private const int RpcCallRejectedRetryDelayMs = 250;
    private const int RPC_E_CALL_REJECTED = unchecked((int)0x80010001);

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
    /// Gets a descriptive string for a slide: "#SlideIndex SlideName".
    /// </summary>
    public static string GetSlideInfo(PowerPoint.Slide slide)
    {
        ArgumentNullException.ThrowIfNull(slide);

        return $"#{slide.SlideIndex} {slide.Name}";
    }

    /// <summary>
    /// Gets the slide index.
    /// </summary>
    public static int GetSlideIndex(PowerPoint.Slide slide)
    {
        ArgumentNullException.ThrowIfNull(slide);

        return slide.SlideIndex;
    }

    /// <summary>
    /// Gets the slide name.
    /// </summary>
    public static string GetSlideName(PowerPoint.Slide slide)
    {
        ArgumentNullException.ThrowIfNull(slide);

        return slide.Name;
    }

    /// <summary>
    /// Sets (renames) the slide name.
    /// </summary>
    public static void SetSlideName(PowerPoint.Slide slide, string name)
    {
        ArgumentNullException.ThrowIfNull(slide);

        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Slide name cannot be empty.", nameof(name));
        }

        slide.Name = name;
    }

    /// <summary>
    /// Gets a descriptive string for a shape: "#ShapeIndex ShapeName".
    /// </summary>
    public static string GetShapeInfo(PowerPoint.Shape shape)
    {
        ArgumentNullException.ThrowIfNull(shape);

        return $"#{shape.Id} {shape.Name}";
    }

    /// <summary>
    /// Gets the formula/text of a shape depending on its type:
    /// - AutoShape, Placeholder, TextBox: reads TextFrame.TextRange.Text
    /// - Picture: reads shape.Title (used as formula source for image shapes)
    /// Returns null if the shape type is not supported or text cannot be read.
    /// Retries on RPC_E_CALL_REJECTED to handle COM busy state.
    /// </summary>
    public static string? GetShapeText(PowerPoint.Shape shape)
    {
        ArgumentNullException.ThrowIfNull(shape);

        if (shape.Type == Core.MsoShapeType.msoPicture)
        {
            string? title = null;

            RetryOnRpcBusy(() =>
            {
                title = shape.Title?.Trim()?.Replace("\r", string.Empty);
            });

            return title;
        }

        if (shape.Type != Core.MsoShapeType.msoAutoShape
            && shape.Type != Core.MsoShapeType.msoPlaceholder
            && shape.Type != Core.MsoShapeType.msoTextBox)
        {
            return null;
        }

        string? result = null;

        RetryOnRpcBusy(() =>
        {
            try
            {
                result = shape.TextFrame?.TextRange?.Text;
            }
            catch (UnauthorizedAccessException)
            {
                result = null;
            }
        });

        return result;
    }

    /// <summary>
    /// Sets the formula/text of a shape depending on its type:
    /// - AutoShape, Placeholder, TextBox: writes TextFrame.TextRange.Text
    /// - Picture: writes shape.Title
    /// Retries on RPC_E_CALL_REJECTED to handle COM busy state.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when the shape type is not supported or has no text frame.</exception>
    public static void SetShapeText(PowerPoint.Shape shape, string text)
    {
        ArgumentNullException.ThrowIfNull(shape);

        if (shape.Type == Core.MsoShapeType.msoPicture)
        {
            RetryOnRpcBusy(() =>
            {
                shape.Title = text;
            });

            return;
        }

        if (shape.Type != Core.MsoShapeType.msoAutoShape
            && shape.Type != Core.MsoShapeType.msoPlaceholder
            && shape.Type != Core.MsoShapeType.msoTextBox)
        {
            throw new InvalidOperationException("The selected shape type does not support text editing.");
        }

        if (shape.HasTextFrame != Core.MsoTriState.msoTrue)
        {
            throw new InvalidOperationException("The selected shape does not have a text frame.");
        }

        RetryOnRpcBusy(() =>
        {
            shape.TextFrame.TextRange.Text = text;
        });
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

    /// <summary>
    /// Retries an action on RPC_E_CALL_REJECTED (0x80010001) which occurs when the COM server is busy.
    /// </summary>
    private static void RetryOnRpcBusy(Action action)
    {
        for (int attempt = 1; attempt <= RpcCallRejectedMaxRetries; attempt++)
        {
            try
            {
                action();
                return;
            }
            catch (COMException ex) when (ex.HResult == RPC_E_CALL_REJECTED)
            {
                if (attempt == RpcCallRejectedMaxRetries)
                {
                    throw;
                }

                Thread.Sleep(RpcCallRejectedRetryDelayMs);
            }
        }
    }
}
