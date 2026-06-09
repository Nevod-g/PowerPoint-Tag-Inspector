namespace PowerPointTagInspector;

/// <summary>
/// Modal dialog for adding or editing a PowerPoint tag.
/// </summary>
internal partial class TagEditorDialog : Form
{
    private readonly bool _isEditMode;

    /// <summary>
    /// Creates a new TagEditorDialog for adding a tag.
    /// </summary>
    public TagEditorDialog()
    {
        InitializeComponent();
        Text = "Add Tag";
        _isEditMode = false;
    }

    /// <summary>
    /// Creates a new TagEditorDialog for editing an existing tag.
    /// </summary>
    public TagEditorDialog(string name, string value)
    {
        InitializeComponent();
        Text = "Edit Tag";
        _isEditMode = true;

        _txtName.Text = name;
        _txtValue.Text = value;
        _txtName.ReadOnly = true;
    }

    /// <summary>
    /// Gets the entered tag name.
    /// </summary>
    public string TagName => _txtName.Text.Trim();

    /// <summary>
    /// Gets the entered tag value.
    /// </summary>
    public string TagValue => _txtValue.Text;

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        base.OnFormClosing(e);

        if (DialogResult != DialogResult.OK)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(_txtName.Text))
        {
            MessageBox.Show(
                this,
                "Tag name cannot be empty.",
                "Validation Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            e.Cancel = true;
            _txtName.Focus();
        }
    }
}
