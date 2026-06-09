namespace PowerPointTagInspector
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            _mainLayout = new System.Windows.Forms.TableLayoutPanel();
            _grpOptions = new System.Windows.Forms.GroupBox();
            _optionsLayout = new System.Windows.Forms.TableLayoutPanel();
            _rbActiveSlide = new System.Windows.Forms.RadioButton();
            _rbSelectedShape = new System.Windows.Forms.RadioButton();
            _grpTags = new System.Windows.Forms.GroupBox();
            _tagsLayout = new System.Windows.Forms.TableLayoutPanel();
            _lvTags = new System.Windows.Forms.ListView();
            _colName = new System.Windows.Forms.ColumnHeader();
            _colValue = new System.Windows.Forms.ColumnHeader();
            _buttonsFlow = new System.Windows.Forms.FlowLayoutPanel();
            _btnAdd = new System.Windows.Forms.Button();
            _btnEdit = new System.Windows.Forms.Button();
            _btnDelete = new System.Windows.Forms.Button();
            _btnRefresh = new System.Windows.Forms.Button();
            _statusStrip = new System.Windows.Forms.StatusStrip();
            _statusLabel = new System.Windows.Forms.ToolStripStatusLabel();

            components = new System.ComponentModel.Container();

            _mainLayout.SuspendLayout();
            _grpOptions.SuspendLayout();
            _optionsLayout.SuspendLayout();
            _grpTags.SuspendLayout();
            _tagsLayout.SuspendLayout();
            _buttonsFlow.SuspendLayout();
            _statusStrip.SuspendLayout();
            SuspendLayout();

            // _mainLayout
            _mainLayout.ColumnCount = 1;
            _mainLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            _mainLayout.RowCount = 2;
            _mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            _mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            _mainLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            _mainLayout.Padding = new System.Windows.Forms.Padding(6);
            _mainLayout.Controls.Add(_grpOptions, 0, 0);
            _mainLayout.Controls.Add(_grpTags, 0, 1);
            _mainLayout.Location = new System.Drawing.Point(0, 0);
            _mainLayout.Name = "_mainLayout";
            _mainLayout.Size = new System.Drawing.Size(600, 380);

            // _grpOptions
            _grpOptions.AutoSize = true;
            _grpOptions.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly;
            _grpOptions.Dock = System.Windows.Forms.DockStyle.Fill;
            _grpOptions.Margin = new System.Windows.Forms.Padding(3);
            _grpOptions.Name = "_grpOptions";
            _grpOptions.Padding = new System.Windows.Forms.Padding(8);
            _grpOptions.Text = "Options";
            _grpOptions.Controls.Add(_optionsLayout);

            // _optionsLayout
            _optionsLayout.ColumnCount = 1;
            _optionsLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            _optionsLayout.RowCount = 2;
            _optionsLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            _optionsLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            _optionsLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            _optionsLayout.Controls.Add(_rbActiveSlide, 0, 0);
            _optionsLayout.Controls.Add(_rbSelectedShape, 0, 1);
            _optionsLayout.Name = "_optionsLayout";

            // _rbActiveSlide
            _rbActiveSlide.AutoSize = true;
            _rbActiveSlide.Checked = true;
            _rbActiveSlide.Margin = new System.Windows.Forms.Padding(3);
            _rbActiveSlide.Name = "_rbActiveSlide";
            _rbActiveSlide.TabIndex = 0;
            _rbActiveSlide.TabStop = true;
            _rbActiveSlide.Text = "Active Slide";
            _rbActiveSlide.AccessibleName = "Active Slide mode";
            _rbActiveSlide.CheckedChanged += RbActiveSlide_CheckedChanged;

            // _rbSelectedShape
            _rbSelectedShape.AutoSize = true;
            _rbSelectedShape.Margin = new System.Windows.Forms.Padding(3);
            _rbSelectedShape.Name = "_rbSelectedShape";
            _rbSelectedShape.TabIndex = 1;
            _rbSelectedShape.Text = "Selected Shape";
            _rbSelectedShape.AccessibleName = "Selected Shape mode";
            _rbSelectedShape.CheckedChanged += RbSelectedShape_CheckedChanged;

            // _grpTags
            _grpTags.Dock = System.Windows.Forms.DockStyle.Fill;
            _grpTags.Margin = new System.Windows.Forms.Padding(3);
            _grpTags.Name = "_grpTags";
            _grpTags.Padding = new System.Windows.Forms.Padding(8);
            _grpTags.Text = "Tags";
            _grpTags.Controls.Add(_tagsLayout);

            // _tagsLayout
            _tagsLayout.ColumnCount = 2;
            _tagsLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            _tagsLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            _tagsLayout.RowCount = 1;
            _tagsLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            _tagsLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            _tagsLayout.Controls.Add(_lvTags, 0, 0);
            _tagsLayout.Controls.Add(_buttonsFlow, 1, 0);
            _tagsLayout.Name = "_tagsLayout";

            // _lvTags
            _lvTags.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] { _colName, _colValue });
            _lvTags.Dock = System.Windows.Forms.DockStyle.Fill;
            _lvTags.FullRowSelect = true;
            _lvTags.GridLines = true;
            _lvTags.Margin = new System.Windows.Forms.Padding(3);
            _lvTags.MultiSelect = false;
            _lvTags.Name = "_lvTags";
            _lvTags.View = System.Windows.Forms.View.Details;
            _lvTags.AccessibleName = "Tags list";
            _lvTags.SelectedIndexChanged += LvTags_SelectedIndexChanged;

            // _colName
            _colName.Text = "Name";
            _colName.Width = 180;

            // _colValue
            _colValue.Text = "Value";
            _colValue.Width = 250;

            // _buttonsFlow
            _buttonsFlow.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            _buttonsFlow.AutoSize = true;
            _buttonsFlow.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            _buttonsFlow.Margin = new System.Windows.Forms.Padding(3);
            _buttonsFlow.Name = "_buttonsFlow";
            _buttonsFlow.Controls.Add(_btnAdd);
            _buttonsFlow.Controls.Add(_btnEdit);
            _buttonsFlow.Controls.Add(_btnDelete);
            _buttonsFlow.Controls.Add(_btnRefresh);

            // _btnAdd
            _btnAdd.Margin = new System.Windows.Forms.Padding(3);
            _btnAdd.Name = "_btnAdd";
            _btnAdd.Size = new System.Drawing.Size(90, 30);
            _btnAdd.Text = "Add...";
            _btnAdd.AccessibleName = "Add Tag";
            _btnAdd.Click += BtnAdd_Click;

            // _btnEdit
            _btnEdit.Enabled = false;
            _btnEdit.Margin = new System.Windows.Forms.Padding(3);
            _btnEdit.Name = "_btnEdit";
            _btnEdit.Size = new System.Drawing.Size(90, 30);
            _btnEdit.Text = "Edit...";
            _btnEdit.AccessibleName = "Edit Tag";
            _btnEdit.Click += BtnEdit_Click;

            // _btnDelete
            _btnDelete.Enabled = false;
            _btnDelete.Margin = new System.Windows.Forms.Padding(3);
            _btnDelete.Name = "_btnDelete";
            _btnDelete.Size = new System.Drawing.Size(90, 30);
            _btnDelete.Text = "Delete";
            _btnDelete.AccessibleName = "Delete Tag";
            _btnDelete.Click += BtnDelete_Click;

            // _btnRefresh
            _btnRefresh.Margin = new System.Windows.Forms.Padding(3, 12, 3, 3);
            _btnRefresh.Name = "_btnRefresh";
            _btnRefresh.Size = new System.Drawing.Size(90, 30);
            _btnRefresh.Text = "Refresh";
            _btnRefresh.AccessibleName = "Refresh Tags";
            _btnRefresh.Click += BtnRefresh_Click;

            // _statusStrip
            _statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { _statusLabel });
            _statusStrip.Name = "_statusStrip";

            // _statusLabel
            _statusLabel.Name = "_statusLabel";
            _statusLabel.Spring = true;
            _statusLabel.Text = "Not connected";
            _statusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;

            // MainForm
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(600, 420);
            Controls.Add(_mainLayout);
            Controls.Add(_statusStrip);
            MinimumSize = new System.Drawing.Size(500, 350);
            Name = "MainForm";
            Text = "PowerPoint Tag Inspector";

            _mainLayout.ResumeLayout(false);
            _mainLayout.PerformLayout();
            _grpOptions.ResumeLayout(false);
            _optionsLayout.ResumeLayout(false);
            _optionsLayout.PerformLayout();
            _grpTags.ResumeLayout(false);
            _tagsLayout.ResumeLayout(false);
            _tagsLayout.PerformLayout();
            _buttonsFlow.ResumeLayout(false);
            _statusStrip.ResumeLayout(false);
            _statusStrip.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel _mainLayout;
        private System.Windows.Forms.GroupBox _grpOptions;
        private System.Windows.Forms.TableLayoutPanel _optionsLayout;
        private System.Windows.Forms.RadioButton _rbActiveSlide;
        private System.Windows.Forms.RadioButton _rbSelectedShape;
        private System.Windows.Forms.GroupBox _grpTags;
        private System.Windows.Forms.TableLayoutPanel _tagsLayout;
        private System.Windows.Forms.ListView _lvTags;
        private System.Windows.Forms.ColumnHeader _colName;
        private System.Windows.Forms.ColumnHeader _colValue;
        private System.Windows.Forms.FlowLayoutPanel _buttonsFlow;
        private System.Windows.Forms.Button _btnAdd;
        private System.Windows.Forms.Button _btnEdit;
        private System.Windows.Forms.Button _btnDelete;
        private System.Windows.Forms.Button _btnRefresh;
        private System.Windows.Forms.StatusStrip _statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel _statusLabel;
    }
}
