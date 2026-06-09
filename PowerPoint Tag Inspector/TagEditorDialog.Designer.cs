namespace PowerPointTagInspector
{
    partial class TagEditorDialog
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
            _tableLayout = new System.Windows.Forms.TableLayoutPanel();
            _lblName = new System.Windows.Forms.Label();
            _txtName = new System.Windows.Forms.TextBox();
            _lblValue = new System.Windows.Forms.Label();
            _txtValue = new System.Windows.Forms.TextBox();
            _flowButtons = new System.Windows.Forms.FlowLayoutPanel();
            _btnCancel = new System.Windows.Forms.Button();
            _btnOK = new System.Windows.Forms.Button();

            _tableLayout.SuspendLayout();
            _flowButtons.SuspendLayout();
            SuspendLayout();

            // _tableLayout
            _tableLayout.ColumnCount = 2;
            _tableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            _tableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            _tableLayout.RowCount = 3;
            _tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            _tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            _tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            _tableLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            _tableLayout.Padding = new System.Windows.Forms.Padding(8);
            _tableLayout.Controls.Add(_lblName, 0, 0);
            _tableLayout.Controls.Add(_txtName, 1, 0);
            _tableLayout.Controls.Add(_lblValue, 0, 1);
            _tableLayout.Controls.Add(_txtValue, 1, 1);
            _tableLayout.Controls.Add(_flowButtons, 0, 2);
            _tableLayout.SetColumnSpan(_flowButtons, 2);
            _tableLayout.Location = new System.Drawing.Point(0, 0);
            _tableLayout.Name = "_tableLayout";
            _tableLayout.Size = new System.Drawing.Size(380, 160);

            // _lblName
            _lblName.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            _lblName.AutoSize = true;
            _lblName.Margin = new System.Windows.Forms.Padding(3, 6, 3, 3);
            _lblName.Name = "_lblName";
            _lblName.Text = "Name:";

            // _txtName
            _txtName.Dock = System.Windows.Forms.DockStyle.Fill;
            _txtName.Margin = new System.Windows.Forms.Padding(3);
            _txtName.Name = "_txtName";
            _txtName.AccessibleName = "Tag Name";

            // _lblValue
            _lblValue.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            _lblValue.AutoSize = true;
            _lblValue.Margin = new System.Windows.Forms.Padding(3, 6, 3, 3);
            _lblValue.Name = "_lblValue";
            _lblValue.Text = "Value:";

            // _txtValue
            _txtValue.Dock = System.Windows.Forms.DockStyle.Fill;
            _txtValue.Margin = new System.Windows.Forms.Padding(3);
            _txtValue.Name = "_txtValue";
            _txtValue.AccessibleName = "Tag Value";

            // _flowButtons
            _flowButtons.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
            _flowButtons.AutoSize = true;
            _flowButtons.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            _flowButtons.Name = "_flowButtons";
            _flowButtons.Controls.Add(_btnCancel);
            _flowButtons.Controls.Add(_btnOK);

            // _btnOK
            _btnOK.Name = "_btnOK";
            _btnOK.Size = new System.Drawing.Size(90, 30);
            _btnOK.Text = "OK";
            _btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            _btnOK.AccessibleName = "OK";

            // _btnCancel
            _btnCancel.Name = "_btnCancel";
            _btnCancel.Size = new System.Drawing.Size(90, 30);
            _btnCancel.Text = "Cancel";
            _btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            _btnCancel.AccessibleName = "Cancel";

            // TagEditorDialog
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(380, 160);
            Controls.Add(_tableLayout);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "TagEditorDialog";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "Tag Editor";
            AcceptButton = _btnOK;
            CancelButton = _btnCancel;

            _tableLayout.ResumeLayout(false);
            _tableLayout.PerformLayout();
            _flowButtons.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel _tableLayout;
        private System.Windows.Forms.Label _lblName;
        private System.Windows.Forms.TextBox _txtName;
        private System.Windows.Forms.Label _lblValue;
        private System.Windows.Forms.TextBox _txtValue;
        private System.Windows.Forms.FlowLayoutPanel _flowButtons;
        private System.Windows.Forms.Button _btnOK;
        private System.Windows.Forms.Button _btnCancel;
    }
}
