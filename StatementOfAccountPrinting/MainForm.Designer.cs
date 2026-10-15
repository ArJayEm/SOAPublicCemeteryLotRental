namespace StatementOfAccountPrinting;

partial class MainForm
{
    private System.ComponentModel.IContainer components = null;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null)) components.Dispose();
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    private void InitializeComponent()
    {
        btnImportExcel = new Button();
        btnGenerateWord = new Button();
        btnPrintPdf = new Button();
        lblStatus = new Label();
        menuStrip1 = new MenuStrip();
        fileToolStripMenuItem = new ToolStripMenuItem();
        importTemplateToolStripMenuItem = new ToolStripMenuItem();
        exportTemplateToolStripMenuItem = new ToolStripMenuItem();
        settingsToolStripMenuItem = new ToolStripMenuItem();
        exitToolStripMenuItem = new ToolStripMenuItem();
        tabBlocks = new TabControl();
        tabPage1 = new TabPage();
        tabPage2 = new TabPage();
        btnBlockFilter = new Button();
        btnPrintSelected = new Button();
        menuStrip1.SuspendLayout();
        tabBlocks.SuspendLayout();
        SuspendLayout();
        // 
        // btnImportExcel
        // 
        btnImportExcel.Location = new Point(12, 26);
        btnImportExcel.Margin = new Padding(3, 2, 3, 2);
        btnImportExcel.Name = "btnImportExcel";
        btnImportExcel.Size = new Size(140, 33);
        btnImportExcel.TabIndex = 0;
        btnImportExcel.Text = "Import Excel";
        btnImportExcel.UseVisualStyleBackColor = true;
        btnImportExcel.Click += btnImportExcel_Click;
        // 
        // btnGenerateWord
        // 
        btnGenerateWord.Enabled = false;
        btnGenerateWord.Location = new Point(158, 26);
        btnGenerateWord.Margin = new Padding(3, 2, 3, 2);
        btnGenerateWord.Name = "btnGenerateWord";
        btnGenerateWord.Size = new Size(140, 33);
        btnGenerateWord.TabIndex = 2;
        btnGenerateWord.Text = "Export as Word (.docx)";
        btnGenerateWord.UseVisualStyleBackColor = true;
        btnGenerateWord.Visible = false;
        btnGenerateWord.Click += btnGenerateWord_Click;
        // 
        // btnPrintPdf
        // 
        btnPrintPdf.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnPrintPdf.Enabled = false;
        btnPrintPdf.Location = new Point(634, 26);
        btnPrintPdf.Margin = new Padding(3, 2, 3, 2);
        btnPrintPdf.Name = "btnPrintPdf";
        btnPrintPdf.Size = new Size(116, 33);
        btnPrintPdf.TabIndex = 3;
        btnPrintPdf.Text = "Export as PDF";
        btnPrintPdf.UseVisualStyleBackColor = true;
        btnPrintPdf.Click += btnPrintPdf_Click;
        // 
        // lblStatus
        // 
        lblStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        lblStatus.Location = new Point(12, 333);
        lblStatus.Name = "lblStatus";
        lblStatus.Size = new Size(721, 19);
        lblStatus.TabIndex = 5;
        lblStatus.Text = "Ready. Import an .xlsx file to begin.";
        // 
        // menuStrip1
        // 
        menuStrip1.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem, settingsToolStripMenuItem, exitToolStripMenuItem });
        menuStrip1.Location = new Point(0, 0);
        menuStrip1.Name = "menuStrip1";
        menuStrip1.Size = new Size(765, 24);
        menuStrip1.TabIndex = 8;
        menuStrip1.Text = "menuStrip1";
        // 
        // fileToolStripMenuItem
        // 
        fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { importTemplateToolStripMenuItem, exportTemplateToolStripMenuItem });
        fileToolStripMenuItem.Name = "fileToolStripMenuItem";
        fileToolStripMenuItem.Size = new Size(37, 20);
        fileToolStripMenuItem.Text = "File";
        // 
        // importTemplateToolStripMenuItem
        // 
        importTemplateToolStripMenuItem.Name = "importTemplateToolStripMenuItem";
        importTemplateToolStripMenuItem.Size = new Size(201, 22);
        importTemplateToolStripMenuItem.Text = "Import Template (.docx)";
        importTemplateToolStripMenuItem.Click += importTemplateToolStripMenuItem_Click;
        // 
        // exportTemplateToolStripMenuItem
        // 
        exportTemplateToolStripMenuItem.Name = "exportTemplateToolStripMenuItem";
        exportTemplateToolStripMenuItem.Size = new Size(201, 22);
        exportTemplateToolStripMenuItem.Text = "Export Template";
        exportTemplateToolStripMenuItem.Click += exportTemplateToolStripMenuItem_Click_1;
        // 
        // settingsToolStripMenuItem
        // 
        settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
        settingsToolStripMenuItem.Size = new Size(61, 20);
        settingsToolStripMenuItem.Text = "Settings";
        settingsToolStripMenuItem.Click += settingsToolStripMenuItem_Click;
        // 
        // exitToolStripMenuItem
        // 
        exitToolStripMenuItem.Name = "exitToolStripMenuItem";
        exitToolStripMenuItem.Size = new Size(38, 20);
        exitToolStripMenuItem.Text = "Exit";
        exitToolStripMenuItem.Click += exitToolStripMenuItem_Click;
        // 
        // tabBlocks
        // 
        tabBlocks.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        tabBlocks.Controls.Add(tabPage1);
        tabBlocks.Controls.Add(tabPage2);
        tabBlocks.Location = new Point(12, 68);
        tabBlocks.Name = "tabBlocks";
        tabBlocks.SelectedIndex = 0;
        tabBlocks.Size = new Size(742, 262);
        tabBlocks.TabIndex = 9;
        // 
        // tabPage1
        // 
        tabPage1.Location = new Point(4, 24);
        tabPage1.Name = "tabPage1";
        tabPage1.Padding = new Padding(3);
        tabPage1.Size = new Size(734, 234);
        tabPage1.TabIndex = 0;
        tabPage1.Text = "tabPage1";
        tabPage1.UseVisualStyleBackColor = true;
        // 
        // tabPage2
        // 
        tabPage2.Location = new Point(4, 24);
        tabPage2.Name = "tabPage2";
        tabPage2.Padding = new Padding(3);
        tabPage2.Size = new Size(734, 234);
        tabPage2.TabIndex = 1;
        tabPage2.Text = "tabPage2";
        tabPage2.UseVisualStyleBackColor = true;
        // 
        // btnBlockFilter
        // 
        btnBlockFilter.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnBlockFilter.Enabled = false;
        btnBlockFilter.Location = new Point(512, 26);
        btnBlockFilter.Margin = new Padding(3, 2, 3, 2);
        btnBlockFilter.Name = "btnBlockFilter";
        btnBlockFilter.Size = new Size(116, 33);
        btnBlockFilter.TabIndex = 13;
        btnBlockFilter.Text = "Blocks (All)";
        btnBlockFilter.UseVisualStyleBackColor = true;
        btnBlockFilter.Click += btnBlockFilter_Click;
        // 
        // btnPrintSelected
        // 
        btnPrintSelected.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnPrintSelected.Enabled = false;
        btnPrintSelected.Location = new Point(374, 26);
        btnPrintSelected.Margin = new Padding(3, 2, 3, 2);
        btnPrintSelected.Name = "btnPrintSelected";
        btnPrintSelected.Size = new Size(116, 33);
        btnPrintSelected.TabIndex = 14;
        btnPrintSelected.Text = "Print Selected";
        btnPrintSelected.UseVisualStyleBackColor = true;
        // 
        // MainForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(765, 361);
        Controls.Add(btnPrintSelected);
        Controls.Add(btnBlockFilter);
        Controls.Add(tabBlocks);
        Controls.Add(lblStatus);
        Controls.Add(btnPrintPdf);
        Controls.Add(btnGenerateWord);
        Controls.Add(btnImportExcel);
        Controls.Add(menuStrip1);
        MainMenuStrip = menuStrip1;
        Margin = new Padding(3, 2, 3, 2);
        MaximizeBox = false;
        MinimumSize = new Size(781, 400);
        Name = "MainForm";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Statement of Account Printing — Public Cemetery Lot Rental";
        WindowState = FormWindowState.Maximized;
        FormClosing += MainForm_FormClosing;
        Shown += MainForm_Shown;
        menuStrip1.ResumeLayout(false);
        menuStrip1.PerformLayout();
        tabBlocks.ResumeLayout(false);
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private System.Windows.Forms.Button btnImportExcel;
    private System.Windows.Forms.Button btnGenerateWord;
    private System.Windows.Forms.Button btnPrintPdf;
    private System.Windows.Forms.Label lblStatus;
    private MenuStrip menuStrip1;
    private ToolStripMenuItem fileToolStripMenuItem;
    private ToolStripMenuItem importTemplateToolStripMenuItem;
    private ToolStripMenuItem exportTemplateToolStripMenuItem;
    private TabControl tabBlocks;
    private TabPage tabPage1;
    private TabPage tabPage2;
    private ToolStripMenuItem settingsToolStripMenuItem;
    private ToolStripMenuItem exitToolStripMenuItem;
    private Button btnBlockFilter;
    private Button btnPrintSelected;
}
