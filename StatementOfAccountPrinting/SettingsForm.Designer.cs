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
            SuspendLayout();
            // 
            // chkSkipZeroTotal
            // 
            chkSkipZeroTotal.AutoSize = true;
            chkSkipZeroTotal.Location = new Point(23, 22);
            chkSkipZeroTotal.Name = "chkSkipZeroTotal";
            chkSkipZeroTotal.Size = new Size(161, 19);
            chkSkipZeroTotal.TabIndex = 0;
            chkSkipZeroTotal.Text = "Skip Zero Total in Printing";
            chkSkipZeroTotal.UseVisualStyleBackColor = true;
            // 
            // cmbPaperSize
            // 
            cmbPaperSize.FormattingEnabled = true;
            cmbPaperSize.Location = new Point(23, 67);
            cmbPaperSize.Name = "cmbPaperSize";
            cmbPaperSize.Size = new Size(140, 23);
            cmbPaperSize.TabIndex = 8;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(22, 49);
            label1.Name = "label1";
            label1.Size = new Size(60, 15);
            label1.TabIndex = 9;
            label1.Text = "Paper Size";
            // 
            // SettingsForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(206, 172);
            Controls.Add(label1);
            Controls.Add(cmbPaperSize);
            Controls.Add(chkSkipZeroTotal);
            FormBorderStyle = FormBorderStyle.None;
            Name = "SettingsForm";
            Text = "SettingsForm";
            FormClosing += SettingsForm_FormClosing;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private CheckBox chkSkipZeroTotal;
        private ComboBox cmbPaperSize;
        private Label label1;
    }
}