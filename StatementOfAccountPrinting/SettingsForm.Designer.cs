namespace StatementOfAccountPrinting
{
    partial class SettingsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            chkSkipZeroTotal = new CheckBox();
            cmbPaperSize = new ComboBox();
            label1 = new Label();
            cmbExportFormat = new ComboBox();
            lblExportFormat = new Label();
            btnSave = new Button();
            btnCancel = new Button();
            groupBox1 = new GroupBox();
            groupBox2 = new GroupBox();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            SuspendLayout();
            // 
            // chkSkipZeroTotal
            // 
            chkSkipZeroTotal.AutoSize = true;
            chkSkipZeroTotal.Location = new Point(9, 22);
            chkSkipZeroTotal.Name = "chkSkipZeroTotal";
            chkSkipZeroTotal.Size = new Size(228, 19);
            chkSkipZeroTotal.TabIndex = 0;
            chkSkipZeroTotal.Text = "Skip records with ₱0.00 total on export";
            chkSkipZeroTotal.UseVisualStyleBackColor = true;
            // 
            // cmbPaperSize
            // 
            cmbPaperSize.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbPaperSize.FormattingEnabled = true;
            cmbPaperSize.Items.AddRange(new object[] { "A4", "Legal" });
            cmbPaperSize.Location = new Point(164, 51);
            cmbPaperSize.Name = "cmbPaperSize";
            cmbPaperSize.Size = new Size(62, 23);
            cmbPaperSize.TabIndex = 8;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(11, 54);
            label1.Name = "label1";
            label1.Size = new Size(86, 15);
            label1.TabIndex = 9;
            label1.Text = "PDF paper size:";
            // 
            // cmbExportFormat
            // 
            cmbExportFormat.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbExportFormat.FormattingEnabled = true;
            cmbExportFormat.Items.AddRange(new object[] { "Docx", "PDF" });
            cmbExportFormat.Location = new Point(164, 22);
            cmbExportFormat.Name = "cmbExportFormat";
            cmbExportFormat.Size = new Size(62, 23);
            cmbExportFormat.TabIndex = 10;
            cmbExportFormat.SelectedIndexChanged += cmbExportFormat_SelectedIndexChanged;
            // 
            // lblExportFormat
            // 
            lblExportFormat.AutoSize = true;
            lblExportFormat.Location = new Point(11, 25);
            lblExportFormat.Name = "lblExportFormat";
            lblExportFormat.Size = new Size(83, 15);
            lblExportFormat.TabIndex = 11;
            lblExportFormat.Text = "Export format:";
            // 
            // btnSave
            // 
            btnSave.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnSave.DialogResult = DialogResult.OK;
            btnSave.FlatStyle = FlatStyle.System;
            btnSave.Location = new Point(176, 177);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(75, 23);
            btnSave.TabIndex = 12;
            btnSave.Text = "Save";
            btnSave.UseVisualStyleBackColor = true;
            btnSave.Click += btnOk_Click;
            // 
            // btnCancel
            // 
            btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.FlatStyle = FlatStyle.System;
            btnCancel.Location = new Point(12, 177);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(75, 23);
            btnCancel.TabIndex = 13;
            btnCancel.Text = "Cancel";
            btnCancel.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            groupBox1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            groupBox1.Controls.Add(lblExportFormat);
            groupBox1.Controls.Add(cmbPaperSize);
            groupBox1.Controls.Add(label1);
            groupBox1.Controls.Add(cmbExportFormat);
            groupBox1.Location = new Point(12, 71);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(243, 85);
            groupBox1.TabIndex = 14;
            groupBox1.TabStop = false;
            groupBox1.Text = "File Format";
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(chkSkipZeroTotal);
            groupBox2.Location = new Point(12, 12);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(243, 48);
            groupBox2.TabIndex = 15;
            groupBox2.TabStop = false;
            groupBox2.Text = "General";
            // 
            // SettingsForm
            // 
            AcceptButton = btnSave;
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            CancelButton = btnCancel;
            ClientSize = new Size(267, 212);
            ControlBox = false;
            Controls.Add(groupBox2);
            Controls.Add(groupBox1);
            Controls.Add(btnCancel);
            Controls.Add(btnSave);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "SettingsForm";
            StartPosition = FormStartPosition.CenterParent;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private CheckBox chkSkipZeroTotal;
        private ComboBox cmbPaperSize;
        private Label label1;
        private ComboBox cmbExportFormat;
        private Label lblExportFormat;
        private Button btnSave;
        private Button btnCancel;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
    }
}