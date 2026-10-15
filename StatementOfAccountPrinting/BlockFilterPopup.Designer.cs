namespace StatementOfAccountPrinting
{
    partial class BlockFilterPopup
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
            btnSelectNone = new Button();
            btnSelectAll = new Button();
            clbBlockFilter = new CheckedListBox();
            SuspendLayout();
            // 
            // btnSelectNone
            // 
            btnSelectNone.Location = new Point(93, 12);
            btnSelectNone.Name = "btnSelectNone";
            btnSelectNone.Size = new Size(81, 23);
            btnSelectNone.TabIndex = 15;
            btnSelectNone.Text = "Select None";
            btnSelectNone.UseVisualStyleBackColor = true;
            btnSelectNone.Click += btnSelectNone_Click;
            // 
            // btnSelectAll
            // 
            btnSelectAll.Location = new Point(12, 12);
            btnSelectAll.Name = "btnSelectAll";
            btnSelectAll.Size = new Size(75, 23);
            btnSelectAll.TabIndex = 14;
            btnSelectAll.Text = "Select All";
            btnSelectAll.UseVisualStyleBackColor = true;
            btnSelectAll.Click += btnSelectAll_Click;
            // 
            // clbBlockFilter
            // 
            clbBlockFilter.CheckOnClick = true;
            clbBlockFilter.FormattingEnabled = true;
            clbBlockFilter.HorizontalScrollbar = true;
            clbBlockFilter.IntegralHeight = false;
            clbBlockFilter.Location = new Point(12, 41);
            clbBlockFilter.Name = "clbBlockFilter";
            clbBlockFilter.Size = new Size(162, 128);
            clbBlockFilter.TabIndex = 13;
            clbBlockFilter.ItemCheck += clbBlockFilter_ItemCheck;
            // 
            // BlockFilterPopup
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(184, 181);
            ControlBox = false;
            Controls.Add(btnSelectNone);
            Controls.Add(btnSelectAll);
            Controls.Add(clbBlockFilter);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "BlockFilterPopup";
            ShowIcon = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            ResumeLayout(false);
        }

        #endregion

        private Button btnSelectNone;
        private Button btnSelectAll;
        private CheckedListBox clbBlockFilter;
    }
}