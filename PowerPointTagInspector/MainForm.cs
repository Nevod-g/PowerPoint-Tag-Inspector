using PowerPointTagInspector.Models;

namespace PowerPointTagInspector;

/// <summary>
/// Main application form for inspecting and editing PowerPoint tags.
/// </summary>
internal partial class MainForm : Form
{
    private readonly TagEditorState _state;

    public MainForm()
    {
        InitializeComponent();
        _state = new TagEditorState();
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        ConnectAndRefresh();
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _state.Dispose();
        base.OnFormClosed(e);
    }

    private void RbActiveSlide_CheckedChanged(object? sender, EventArgs e)
    {
        if (_rbActiveSlide.Checked)
        {
            _state.SetMode(TagSourceMode.ActiveSlide);
            UpdateUI();
        }
    }

    private void RbSelectedShape_CheckedChanged(object? sender, EventArgs e)
    {
        if (_rbSelectedShape.Checked)
        {
            _state.SetMode(TagSourceMode.SelectedShape);
            UpdateUI();
        }
    }

    private void BtnAdd_Click(object? sender, EventArgs e)
    {
        using TagEditorDialog dialog = new();

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        string tagName = dialog.TagName;
        string tagValue = dialog.TagValue;

        if (_state.TagExists(tagName))
        {
            DialogResult confirm = MessageBox.Show(
                this,
                $"A tag with name '{tagName}' already exists. Overwrite its value?",
                "Tag Exists",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm != DialogResult.Yes)
            {
                return;
            }
        }

        try
        {
            _state.SetTag(tagName, tagValue);
            _state.Refresh();
            UpdateUI();
        }
        catch (InvalidOperationException ex)
        {
            ShowError("Failed to add tag.", ex.Message);
        }
    }

    private void BtnEdit_Click(object? sender, EventArgs e)
    {
        if (_lvTags.SelectedItems.Count == 0)
        {
            return;
        }

        ListViewItem selectedItem = _lvTags.SelectedItems[0];
        string currentName = selectedItem.Text;
        string currentValue = selectedItem.SubItems[1].Text;

        using TagEditorDialog dialog = new(currentName, currentValue);

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        try
        {
            _state.SetTag(dialog.TagName, dialog.TagValue);
            _state.Refresh();
            UpdateUI();
        }
        catch (InvalidOperationException ex)
        {
            ShowError("Failed to edit tag.", ex.Message);
        }
    }

    private void BtnDelete_Click(object? sender, EventArgs e)
    {
        if (_lvTags.SelectedItems.Count == 0)
        {
            return;
        }

        string tagName = _lvTags.SelectedItems[0].Text;

        DialogResult confirm = MessageBox.Show(
            this,
            $"Are you sure you want to delete tag '{tagName}'?",
            "Confirm Delete",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (confirm != DialogResult.Yes)
        {
            return;
        }

        try
        {
            _state.RemoveTag(tagName);
            _state.Refresh();
            UpdateUI();
        }
        catch (InvalidOperationException ex)
        {
            ShowError("Failed to delete tag.", ex.Message);
        }
    }

    private void BtnRefresh_Click(object? sender, EventArgs e)
    {
        if (!_state.IsConnected)
        {
            ConnectAndRefresh();
        }
        else
        {
            _state.Refresh();
            UpdateUI();
        }
    }

    private void LvTags_SelectedIndexChanged(object? sender, EventArgs e)
    {
        bool hasSelection = _lvTags.SelectedItems.Count > 0;
        _btnEdit.Enabled = hasSelection;
        _btnDelete.Enabled = hasSelection;
    }

    private void ConnectAndRefresh()
    {
        try
        {
            _state.Connect();
            _state.Refresh();
        }
        catch (InvalidOperationException)
        {
            // Error is captured in _state.LastError
        }

        UpdateUI();
    }

    private void UpdateUI()
    {
        _lvTags.BeginUpdate();
        _lvTags.Items.Clear();

        foreach (TagItem tag in _state.CurrentTags)
        {
            if (tag.Name == "TAG") continue; // skip internal tag
            _lvTags.Items.Add(new ListViewItem([tag.Name, tag.Value]));
        }

        _lvTags.EndUpdate();

        // Auto-size columns to content
        _colName.Width = -2;
        _colValue.Width = -2;

        // Ensure minimum column widths
        if (_colName.Width < 150)
        {
            _colName.Width = 150;
        }

        if (_colValue.Width < 200)
        {
            _colValue.Width = 200;
        }

        bool hasTarget = _state.IsConnected && _state.LastError is null;
        _btnAdd.Enabled = hasTarget;
        _btnEdit.Enabled = hasTarget && _lvTags.SelectedItems.Count > 0;
        _btnDelete.Enabled = hasTarget && _lvTags.SelectedItems.Count > 0;

        // Update radio button labels with current target info
        UpdateRadioButtonLabels();

        // Update slide name editor
        UpdateSlideNameEditor();

        // Update shape text editor
        UpdateShapeTextEditor();

        _statusLabel.Text = _state.GetStatusText();
    }

    private void UpdateSlideNameEditor()
    {
        if (_state.SlideIndex is not null)
        {
            _lblSlideIndex.Text = $"Slide #{_state.SlideIndex}";
            _txtSlideName.Text = _state.SlideName ?? string.Empty;
            _txtSlideName.Enabled = true;
            _btnApplySlideName.Enabled = true;
        }
        else
        {
            _lblSlideIndex.Text = "Slide #";
            _txtSlideName.Text = string.Empty;
            _txtSlideName.Enabled = false;
            _btnApplySlideName.Enabled = false;
        }
    }

    private void BtnApplySlideName_Click(object? sender, EventArgs e)
    {
        string newName = _txtSlideName.Text.Trim();

        if (string.IsNullOrWhiteSpace(newName))
        {
            MessageBox.Show(this, "Slide name cannot be empty.", "Validation Error",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            _txtSlideName.Focus();
            return;
        }

        try
        {
            _state.RenameSlideName(newName);
            _state.Refresh();
            UpdateUI();
        }
        catch (InvalidOperationException ex)
        {
            ShowError("Failed to rename slide.", ex.Message);
        }
    }

    private void UpdateShapeTextEditor()
    {
        bool isShapeMode = _state.CurrentMode == TagSourceMode.SelectedShape
                           && _state.LastError is null;

        if (isShapeMode)
        {
            _lblShapeText.Text = "Formula:";
            _txtShapeText.Text = _state.ShapeText ?? string.Empty;
            _txtShapeText.Enabled = _state.ShapeText is not null;
            _btnApplyShapeText.Enabled = _state.ShapeText is not null;
        }
        else
        {
            _lblShapeText.Text = "Formula:";
            _txtShapeText.Text = string.Empty;
            _txtShapeText.Enabled = false;
            _btnApplyShapeText.Enabled = false;
        }
    }

    private void BtnApplyShapeText_Click(object? sender, EventArgs e)
    {
        string newText = _txtShapeText.Text;

        try
        {
            _state.SetShapeText(newText);
            _state.Refresh();
            UpdateUI();
        }
        catch (InvalidOperationException ex)
        {
            ShowError("Failed to update shape text.", ex.Message);
        }
    }

    private void UpdateRadioButtonLabels()
    {
        string slideLabel = "Active Slide";
        string shapeLabel = "Selected Shape";

        if (_state.IsConnected && _state.TargetDescription is not null)
        {
            if (_state.CurrentMode == TagSourceMode.ActiveSlide)
            {
                slideLabel = $"Active Slide ({_state.TargetDescription})";
            }
            else
            {
                shapeLabel = $"Selected Shape ({_state.TargetDescription})";
            }
        }

        _rbActiveSlide.Text = slideLabel;
        _rbSelectedShape.Text = shapeLabel;
    }

    private void ShowError(string title, string message)
    {
        MessageBox.Show(this, message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        _state.Refresh();
        UpdateUI();
    }
}
