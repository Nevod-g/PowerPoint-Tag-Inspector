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
            _mainLayout = new TableLayoutPanel();
            _grpOptions = new GroupBox();
            _optionsLayout = new TableLayoutPanel();
            _rbActiveSlide = new RadioButton();
            _rbSelectedShape = new RadioButton();
            _slideNameLayout = new TableLayoutPanel();
            _lblSlideIndex = new Label();
            _txtSlideName = new TextBox();
            _btnApplySlideName = new Button();
            _grpTags = new GroupBox();
            _tagsLayout = new TableLayoutPanel();
            _lvTags = new ListView();
            _colName = new ColumnHeader();
            _colValue = new ColumnHeader();
            _buttonsFlow = new FlowLayoutPanel();
            _btnAdd = new Button();
            _btnEdit = new Button();
            _btnDelete = new Button();
            _btnRefresh = new Button();
            _statusStrip = new StatusStrip();
            _statusLabel = new ToolStripStatusLabel();
            _mainLayout.SuspendLayout();
            _grpOptions.SuspendLayout();
            _optionsLayout.SuspendLayout();
            _slideNameLayout.SuspendLayout();
            _grpTags.SuspendLayout();
            _tagsLayout.SuspendLayout();
            _buttonsFlow.SuspendLayout();
            _statusStrip.SuspendLayout();
            SuspendLayout();
            // 
            // _mainLayout
            // 
            _mainLayout.ColumnCount = 1;
            _mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _mainLayout.Controls.Add(_grpOptions, 0, 0);
            _mainLayout.Controls.Add(_slideNameLayout, 0, 1);
            _mainLayout.Controls.Add(_grpTags, 0, 2);
            _mainLayout.Dock = DockStyle.Fill;
            _mainLayout.Location = new Point(0, 0);
            _mainLayout.Name = "_mainLayout";
            _mainLayout.Padding = new Padding(6);
            _mainLayout.RowCount = 3;
            _mainLayout.RowStyles.Add(new RowStyle());
            _mainLayout.RowStyles.Add(new RowStyle());
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            _mainLayout.Size = new Size(923, 398);
            _mainLayout.TabIndex = 0;
            // 
            // _grpOptions
            // 
            _grpOptions.AutoSize = true;
            _grpOptions.Controls.Add(_optionsLayout);
            _grpOptions.Dock = DockStyle.Fill;
            _grpOptions.Location = new Point(9, 9);
            _grpOptions.Name = "_grpOptions";
            _grpOptions.Padding = new Padding(8);
            _grpOptions.Size = new Size(905, 82);
            _grpOptions.TabIndex = 0;
            _grpOptions.TabStop = false;
            _grpOptions.Text = "Options";
            // 
            // _optionsLayout
            // 
            _optionsLayout.AutoSize = true;
            _optionsLayout.ColumnCount = 1;
            _optionsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _optionsLayout.Controls.Add(_rbActiveSlide, 0, 0);
            _optionsLayout.Controls.Add(_rbSelectedShape, 0, 1);
            _optionsLayout.Dock = DockStyle.Top;
            _optionsLayout.Location = new Point(8, 24);
            _optionsLayout.Name = "_optionsLayout";
            _optionsLayout.RowCount = 2;
            _optionsLayout.RowStyles.Add(new RowStyle());
            _optionsLayout.RowStyles.Add(new RowStyle());
            _optionsLayout.Size = new Size(889, 50);
            _optionsLayout.TabIndex = 0;
            // 
            // _rbActiveSlide
            // 
            _rbActiveSlide.AccessibleName = "Active Slide mode";
            _rbActiveSlide.AutoSize = true;
            _rbActiveSlide.Checked = true;
            _rbActiveSlide.Location = new Point(3, 3);
            _rbActiveSlide.Name = "_rbActiveSlide";
            _rbActiveSlide.Size = new Size(86, 19);
            _rbActiveSlide.TabIndex = 0;
            _rbActiveSlide.TabStop = true;
            _rbActiveSlide.Text = "Active Slide";
            _rbActiveSlide.CheckedChanged += RbActiveSlide_CheckedChanged;
            // 
            // _rbSelectedShape
            // 
            _rbSelectedShape.AccessibleName = "Selected Shape mode";
            _rbSelectedShape.AutoSize = true;
            _rbSelectedShape.Location = new Point(3, 28);
            _rbSelectedShape.Name = "_rbSelectedShape";
            _rbSelectedShape.Size = new Size(104, 19);
            _rbSelectedShape.TabIndex = 1;
            _rbSelectedShape.Text = "Selected Shape";
            _rbSelectedShape.CheckedChanged += RbSelectedShape_CheckedChanged;
            // 
            // _slideNameLayout
            // 
            _slideNameLayout.AutoSize = true;
            _slideNameLayout.ColumnCount = 3;
            _slideNameLayout.ColumnStyles.Add(new ColumnStyle());
            _slideNameLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _slideNameLayout.ColumnStyles.Add(new ColumnStyle());
            _slideNameLayout.Controls.Add(_lblSlideIndex, 0, 0);
            _slideNameLayout.Controls.Add(_txtSlideName, 1, 0);
            _slideNameLayout.Controls.Add(_btnApplySlideName, 2, 0);
            _slideNameLayout.Dock = DockStyle.Fill;
            _slideNameLayout.Margin = new Padding(3, 0, 3, 0);
            _slideNameLayout.Name = "_slideNameLayout";
            _slideNameLayout.RowCount = 1;
            _slideNameLayout.RowStyles.Add(new RowStyle());
            _slideNameLayout.Size = new Size(905, 30);
            _slideNameLayout.TabIndex = 1;
            // 
            // _lblSlideIndex
            // 
            _lblSlideIndex.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            _lblSlideIndex.AutoSize = true;
            _lblSlideIndex.Location = new Point(3, 7);
            _lblSlideIndex.Margin = new Padding(3);
            _lblSlideIndex.Name = "_lblSlideIndex";
            _lblSlideIndex.Text = "Slide #";
            _lblSlideIndex.AccessibleName = "Slide index";
            // 
            // _txtSlideName
            // 
            _txtSlideName.Dock = DockStyle.Fill;
            _txtSlideName.Margin = new Padding(3);
            _txtSlideName.Name = "_txtSlideName";
            _txtSlideName.AccessibleName = "Slide Name";
            // 
            // _btnApplySlideName
            // 
            _btnApplySlideName.Location = new Point(3, 3);
            _btnApplySlideName.Margin = new Padding(3);
            _btnApplySlideName.Name = "_btnApplySlideName";
            _btnApplySlideName.Size = new Size(75, 23);
            _btnApplySlideName.Text = "Apply";
            _btnApplySlideName.AccessibleName = "Apply slide name";
            _btnApplySlideName.Click += BtnApplySlideName_Click;
            // 
            // _grpTags
            // 
            _grpTags.Controls.Add(_tagsLayout);
            _grpTags.Dock = DockStyle.Fill;
            _grpTags.Location = new Point(9, 97);
            _grpTags.Name = "_grpTags";
            _grpTags.Padding = new Padding(8);
            _grpTags.Size = new Size(905, 292);
            _grpTags.TabIndex = 1;
            _grpTags.TabStop = false;
            _grpTags.Text = "Tags";
            // 
            // _tagsLayout
            // 
            _tagsLayout.ColumnCount = 2;
            _tagsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _tagsLayout.ColumnStyles.Add(new ColumnStyle());
            _tagsLayout.Controls.Add(_lvTags, 0, 0);
            _tagsLayout.Controls.Add(_buttonsFlow, 1, 0);
            _tagsLayout.Dock = DockStyle.Fill;
            _tagsLayout.Location = new Point(8, 24);
            _tagsLayout.Name = "_tagsLayout";
            _tagsLayout.RowCount = 1;
            _tagsLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            _tagsLayout.Size = new Size(889, 260);
            _tagsLayout.TabIndex = 0;
            // 
            // _lvTags
            // 
            _lvTags.AccessibleName = "Tags list";
            _lvTags.Columns.AddRange(new ColumnHeader[] { _colName, _colValue });
            _lvTags.Dock = DockStyle.Fill;
            _lvTags.FullRowSelect = true;
            _lvTags.GridLines = true;
            _lvTags.Location = new Point(3, 3);
            _lvTags.MultiSelect = false;
            _lvTags.Name = "_lvTags";
            _lvTags.Size = new Size(781, 254);
            _lvTags.TabIndex = 0;
            _lvTags.UseCompatibleStateImageBehavior = false;
            _lvTags.View = View.Details;
            _lvTags.SelectedIndexChanged += LvTags_SelectedIndexChanged;
            // 
            // _colName
            // 
            _colName.Text = "Name";
            _colName.Width = 180;
            // 
            // _colValue
            // 
            _colValue.Text = "Value";
            _colValue.Width = 250;
            // 
            // _buttonsFlow
            // 
            _buttonsFlow.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _buttonsFlow.AutoSize = true;
            _buttonsFlow.Controls.Add(_btnAdd);
            _buttonsFlow.Controls.Add(_btnEdit);
            _buttonsFlow.Controls.Add(_btnDelete);
            _buttonsFlow.Controls.Add(_btnRefresh);
            _buttonsFlow.FlowDirection = FlowDirection.TopDown;
            _buttonsFlow.Location = new Point(790, 3);
            _buttonsFlow.Name = "_buttonsFlow";
            _buttonsFlow.Size = new Size(96, 153);
            _buttonsFlow.TabIndex = 1;
            // 
            // _btnAdd
            // 
            _btnAdd.AccessibleName = "Add Tag";
            _btnAdd.Location = new Point(3, 3);
            _btnAdd.Name = "_btnAdd";
            _btnAdd.Size = new Size(90, 30);
            _btnAdd.TabIndex = 0;
            _btnAdd.Text = "Add...";
            _btnAdd.Click += BtnAdd_Click;
            // 
            // _btnEdit
            // 
            _btnEdit.AccessibleName = "Edit Tag";
            _btnEdit.Enabled = false;
            _btnEdit.Location = new Point(3, 39);
            _btnEdit.Name = "_btnEdit";
            _btnEdit.Size = new Size(90, 30);
            _btnEdit.TabIndex = 1;
            _btnEdit.Text = "Edit...";
            _btnEdit.Click += BtnEdit_Click;
            // 
            // _btnDelete
            // 
            _btnDelete.AccessibleName = "Delete Tag";
            _btnDelete.Enabled = false;
            _btnDelete.Location = new Point(3, 75);
            _btnDelete.Name = "_btnDelete";
            _btnDelete.Size = new Size(90, 30);
            _btnDelete.TabIndex = 2;
            _btnDelete.Text = "Delete";
            _btnDelete.Click += BtnDelete_Click;
            // 
            // _btnRefresh
            // 
            _btnRefresh.AccessibleName = "Refresh Tags";
            _btnRefresh.Location = new Point(3, 120);
            _btnRefresh.Margin = new Padding(3, 12, 3, 3);
            _btnRefresh.Name = "_btnRefresh";
            _btnRefresh.Size = new Size(90, 30);
            _btnRefresh.TabIndex = 3;
            _btnRefresh.Text = "Refresh";
            _btnRefresh.Click += BtnRefresh_Click;
            // 
            // _statusStrip
            // 
            _statusStrip.Items.AddRange(new ToolStripItem[] { _statusLabel });
            _statusStrip.Location = new Point(0, 398);
            _statusStrip.Name = "_statusStrip";
            _statusStrip.Size = new Size(923, 22);
            _statusStrip.TabIndex = 1;
            // 
            // _statusLabel
            // 
            _statusLabel.Name = "_statusLabel";
            _statusLabel.Size = new Size(908, 17);
            _statusLabel.Spring = true;
            _statusLabel.Text = "Not connected";
            _statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(923, 420);
            Controls.Add(_mainLayout);
            Controls.Add(_statusStrip);
            MinimumSize = new Size(500, 350);
            Name = "MainForm";
            Text = "PowerPoint Tag Inspector";
            _mainLayout.ResumeLayout(false);
            _mainLayout.PerformLayout();
            _grpOptions.ResumeLayout(false);
            _grpOptions.PerformLayout();
            _optionsLayout.ResumeLayout(false);
            _optionsLayout.PerformLayout();
            _slideNameLayout.ResumeLayout(false);
            _slideNameLayout.PerformLayout();
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
        private System.Windows.Forms.TableLayoutPanel _slideNameLayout;
        private System.Windows.Forms.Label _lblSlideIndex;
        private System.Windows.Forms.TextBox _txtSlideName;
        private System.Windows.Forms.Button _btnApplySlideName;
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
