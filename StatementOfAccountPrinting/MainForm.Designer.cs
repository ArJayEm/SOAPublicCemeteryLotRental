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
        System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
        lblStatus = new Label();
        menuStrip1 = new MenuStrip();
        fileToolStripMenuItem = new ToolStripMenuItem();
        importTemplateToolStripMenuItem = new ToolStripMenuItem();
        exportTemplateToolStripMenuItem = new ToolStripMenuItem();
        toolStripSeparator2 = new ToolStripSeparator();
        exitToolStripMenuItem = new ToolStripMenuItem();
        settingsToolStripMenuItem = new ToolStripMenuItem();
        helpToolStripMenuItem = new ToolStripMenuItem();
        checkForUpdatesToolStripMenuItem = new ToolStripMenuItem();
        toolStripSeparator1 = new ToolStripSeparator();
        aboutToolStripMenuItem = new ToolStripMenuItem();
        toolStrip1 = new ToolStrip();
        toolStripSplitButtonExport = new ToolStripSplitButton();
        toolStripMenuItemExportSelected = new ToolStripMenuItem();
        exportAllskip000TotalAmountDueToolStripMenuItem = new ToolStripMenuItem();
        exportAllonly000TotalAmountDueToolStripMenuItem = new ToolStripMenuItem();
        toolStripButtonImportExcel = new ToolStripButton();
        toolStripMenuItemBlockFilter = new ToolStripButton();
        toolStripTextBoxSearch = new ToolStripTextBox();
        toolStripButtonClear = new ToolStripButton();
        tabPage2 = new TabPage();
        tabPage1 = new TabPage();
        tabBlocks = new TabControl();
        menuStrip1.SuspendLayout();
        toolStrip1.SuspendLayout();
        tabBlocks.SuspendLayout();
        SuspendLayout();
        // 
        // lblStatus
        // 
        lblStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        lblStatus.Location = new Point(12, 347);
        lblStatus.Name = "lblStatus";
        lblStatus.Size = new Size(721, 19);
        lblStatus.TabIndex = 5;
        lblStatus.Text = "Ready. Import an .xlsx file to begin.";
        // 
        // menuStrip1
        // 
        menuStrip1.Font = new Font("Segoe UI", 12F);
        menuStrip1.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem, settingsToolStripMenuItem, helpToolStripMenuItem });
        menuStrip1.Location = new Point(0, 0);
        menuStrip1.Name = "menuStrip1";
        menuStrip1.Size = new Size(765, 29);
        menuStrip1.TabIndex = 8;
        menuStrip1.Text = "menuStrip1";
        // 
        // fileToolStripMenuItem
        // 
        fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { importTemplateToolStripMenuItem, exportTemplateToolStripMenuItem, toolStripSeparator2, exitToolStripMenuItem });
        fileToolStripMenuItem.Name = "fileToolStripMenuItem";
        fileToolStripMenuItem.ShortcutKeys = Keys.Alt | Keys.F;
        fileToolStripMenuItem.Size = new Size(46, 25);
        fileToolStripMenuItem.Text = "File";
        // 
        // importTemplateToolStripMenuItem
        // 
        importTemplateToolStripMenuItem.Name = "importTemplateToolStripMenuItem";
        importTemplateToolStripMenuItem.ShortcutKeys = Keys.Alt | Keys.I;
        importTemplateToolStripMenuItem.Size = new Size(286, 26);
        importTemplateToolStripMenuItem.Text = "Import Template (.docx)";
        importTemplateToolStripMenuItem.Click += importTemplateToolStripMenuItem_Click;
        // 
        // exportTemplateToolStripMenuItem
        // 
        exportTemplateToolStripMenuItem.Name = "exportTemplateToolStripMenuItem";
        exportTemplateToolStripMenuItem.ShortcutKeys = Keys.Alt | Keys.E;
        exportTemplateToolStripMenuItem.Size = new Size(286, 26);
        exportTemplateToolStripMenuItem.Text = "Export Template";
        exportTemplateToolStripMenuItem.Click += exportTemplateToolStripMenuItem_Click_1;
        // 
        // toolStripSeparator2
        // 
        toolStripSeparator2.Name = "toolStripSeparator2";
        toolStripSeparator2.Size = new Size(283, 6);
        // 
        // exitToolStripMenuItem
        // 
        exitToolStripMenuItem.Name = "exitToolStripMenuItem";
        exitToolStripMenuItem.ShortcutKeys = Keys.Alt | Keys.F4;
        exitToolStripMenuItem.Size = new Size(286, 26);
        exitToolStripMenuItem.Text = "Exit";
        exitToolStripMenuItem.Click += exitToolStripMenuItem_Click;
        // 
        // settingsToolStripMenuItem
        // 
        settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
        settingsToolStripMenuItem.ShortcutKeys = Keys.Alt | Keys.S;
        settingsToolStripMenuItem.Size = new Size(78, 25);
        settingsToolStripMenuItem.Text = "Settings";
        settingsToolStripMenuItem.Click += settingsToolStripMenuItem_Click;
        // 
        // helpToolStripMenuItem
        // 
        helpToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { checkForUpdatesToolStripMenuItem, toolStripSeparator1, aboutToolStripMenuItem });
        helpToolStripMenuItem.Name = "helpToolStripMenuItem";
        helpToolStripMenuItem.ShortcutKeys = Keys.Alt | Keys.H;
        helpToolStripMenuItem.Size = new Size(54, 25);
        helpToolStripMenuItem.Text = "Help";
        // 
        // checkForUpdatesToolStripMenuItem
        // 
        checkForUpdatesToolStripMenuItem.Name = "checkForUpdatesToolStripMenuItem";
        checkForUpdatesToolStripMenuItem.ShortcutKeys = Keys.Alt | Keys.U;
        checkForUpdatesToolStripMenuItem.Size = new Size(261, 26);
        checkForUpdatesToolStripMenuItem.Text = "Check For Updates";
        checkForUpdatesToolStripMenuItem.Click += checkForUpdatesToolStripMenuItem_Click;
        // 
        // toolStripSeparator1
        // 
        toolStripSeparator1.Name = "toolStripSeparator1";
        toolStripSeparator1.Size = new Size(258, 6);
        // 
        // aboutToolStripMenuItem
        // 
        aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
        aboutToolStripMenuItem.ShortcutKeys = Keys.Alt | Keys.A;
        aboutToolStripMenuItem.Size = new Size(261, 26);
        aboutToolStripMenuItem.Text = "About";
        aboutToolStripMenuItem.Click += aboutToolStripMenuItem_Click;
        // 
        // toolStrip1
        // 
        toolStrip1.Font = new Font("Segoe UI", 12F);
        toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripSplitButtonExport, toolStripButtonImportExcel, toolStripMenuItemBlockFilter, toolStripTextBoxSearch, toolStripButtonClear });
        toolStrip1.Location = new Point(0, 29);
        toolStrip1.Name = "toolStrip1";
        toolStrip1.Padding = new Padding(10, 0, 10, 0);
        toolStrip1.Size = new Size(765, 28);
        toolStrip1.TabIndex = 15;
        toolStrip1.Text = "toolStrip1";
        // 
        // toolStripSplitButtonExport
        // 
        toolStripSplitButtonExport.Alignment = ToolStripItemAlignment.Right;
        toolStripSplitButtonExport.DropDownButtonWidth = 15;
        toolStripSplitButtonExport.DropDownItems.AddRange(new ToolStripItem[] { toolStripMenuItemExportSelected, exportAllskip000TotalAmountDueToolStripMenuItem, exportAllonly000TotalAmountDueToolStripMenuItem });
        toolStripSplitButtonExport.Font = new Font("Segoe UI", 12F);
        toolStripSplitButtonExport.Image = Properties.Resources.pdf;
        toolStripSplitButtonExport.ImageTransparentColor = Color.Magenta;
        toolStripSplitButtonExport.Name = "toolStripSplitButtonExport";
        toolStripSplitButtonExport.Size = new Size(90, 25);
        toolStripSplitButtonExport.Text = "Export";
        toolStripSplitButtonExport.ButtonClick += toolStripSplitButtonExport_ButtonClick;
        toolStripSplitButtonExport.Click += toolStripSplitButtonExport_Click;
        // 
        // toolStripMenuItemExportSelected
        // 
        toolStripMenuItemExportSelected.Name = "toolStripMenuItemExportSelected";
        toolStripMenuItemExportSelected.Size = new Size(361, 26);
        toolStripMenuItemExportSelected.Text = "Export Selected Record(s)";
        toolStripMenuItemExportSelected.Click += toolStripMenuItemExportSelected_Click;
        // 
        // exportAllskip000TotalAmountDueToolStripMenuItem
        // 
        exportAllskip000TotalAmountDueToolStripMenuItem.Name = "exportAllskip000TotalAmountDueToolStripMenuItem";
        exportAllskip000TotalAmountDueToolStripMenuItem.Size = new Size(361, 26);
        exportAllskip000TotalAmountDueToolStripMenuItem.Text = "Export All (skip ₱0.00 Total Amount Due)";
        exportAllskip000TotalAmountDueToolStripMenuItem.Click += exportAllskip000TotalAmountDueToolStripMenuItem_Click;
        // 
        // exportAllonly000TotalAmountDueToolStripMenuItem
        // 
        exportAllonly000TotalAmountDueToolStripMenuItem.Name = "exportAllonly000TotalAmountDueToolStripMenuItem";
        exportAllonly000TotalAmountDueToolStripMenuItem.Size = new Size(361, 26);
        exportAllonly000TotalAmountDueToolStripMenuItem.Text = "Export All (only ₱0.00 Total Amount Due)";
        exportAllonly000TotalAmountDueToolStripMenuItem.Visible = false;
        exportAllonly000TotalAmountDueToolStripMenuItem.Click += exportAllonly000TotalAmountDueToolStripMenuItem_Click;
        // 
        // toolStripButtonImportExcel
        // 
        toolStripButtonImportExcel.Image = (Image)resources.GetObject("toolStripButtonImportExcel.Image");
        toolStripButtonImportExcel.ImageAlign = ContentAlignment.MiddleLeft;
        toolStripButtonImportExcel.ImageTransparentColor = Color.Magenta;
        toolStripButtonImportExcel.Name = "toolStripButtonImportExcel";
        toolStripButtonImportExcel.Size = new Size(115, 25);
        toolStripButtonImportExcel.Text = "Import Excel";
        toolStripButtonImportExcel.Click += toolStripButtonImportExcel_Click;
        // 
        // toolStripMenuItemBlockFilter
        // 
        toolStripMenuItemBlockFilter.Alignment = ToolStripItemAlignment.Right;
        toolStripMenuItemBlockFilter.DisplayStyle = ToolStripItemDisplayStyle.Text;
        toolStripMenuItemBlockFilter.Image = (Image)resources.GetObject("toolStripMenuItemBlockFilter.Image");
        toolStripMenuItemBlockFilter.ImageTransparentColor = Color.Magenta;
        toolStripMenuItemBlockFilter.Name = "toolStripMenuItemBlockFilter";
        toolStripMenuItemBlockFilter.Size = new Size(101, 25);
        toolStripMenuItemBlockFilter.Text = "All Blocks (*)";
        toolStripMenuItemBlockFilter.Click += toolStripMenuItemBlockFilter_Click;
        // 
        // toolStripTextBoxSearch
        // 
        toolStripTextBoxSearch.Name = "toolStripTextBoxSearch";
        toolStripTextBoxSearch.Size = new Size(200, 28);
        toolStripTextBoxSearch.ToolTipText = "Search List";
        toolStripTextBoxSearch.KeyDown += toolStripTextBoxSearch_KeyDown;
        toolStripTextBoxSearch.TextChanged += toolStripTextBoxSearch_TextChanged;
        // 
        // toolStripButtonClear
        // 
        toolStripButtonClear.DisplayStyle = ToolStripItemDisplayStyle.Image;
        toolStripButtonClear.Enabled = false;
        toolStripButtonClear.Image = Properties.Resources.close;
        toolStripButtonClear.ImageTransparentColor = Color.Magenta;
        toolStripButtonClear.Name = "toolStripButtonClear";
        toolStripButtonClear.Size = new Size(23, 25);
        toolStripButtonClear.Click += toolStripButtonClear_Click;
        // 
        // tabPage2
        // 
        tabPage2.Location = new Point(4, 26);
        tabPage2.Name = "tabPage2";
        tabPage2.Padding = new Padding(3);
        tabPage2.Size = new Size(734, 254);
        tabPage2.TabIndex = 1;
        tabPage2.Text = "tabPage2";
        tabPage2.UseVisualStyleBackColor = true;
        // 
        // tabPage1
        // 
        tabPage1.Location = new Point(4, 26);
        tabPage1.Name = "tabPage1";
        tabPage1.Padding = new Padding(3);
        tabPage1.Size = new Size(734, 254);
        tabPage1.TabIndex = 0;
        tabPage1.Text = "tabPage1";
        tabPage1.UseVisualStyleBackColor = true;
        // 
        // tabBlocks
        // 
        tabBlocks.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        tabBlocks.Controls.Add(tabPage1);
        tabBlocks.Controls.Add(tabPage2);
        tabBlocks.Font = new Font("Segoe UI", 10F);
        tabBlocks.Location = new Point(12, 60);
        tabBlocks.Name = "tabBlocks";
        tabBlocks.SelectedIndex = 0;
        tabBlocks.Size = new Size(742, 284);
        tabBlocks.TabIndex = 9;
        // 
        // MainForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(765, 375);
        Controls.Add(toolStrip1);
        Controls.Add(tabBlocks);
        Controls.Add(lblStatus);
        Controls.Add(menuStrip1);
        MainMenuStrip = menuStrip1;
        Margin = new Padding(3, 2, 3, 2);
        MinimumSize = new Size(781, 400);
        Name = "MainForm";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "SoA Printing — Public Cemetery Lot Rental";
        WindowState = FormWindowState.Maximized;
        FormClosing += MainForm_FormClosing;
        Shown += MainForm_Shown;
        menuStrip1.ResumeLayout(false);
        menuStrip1.PerformLayout();
        toolStrip1.ResumeLayout(false);
        toolStrip1.PerformLayout();
        tabBlocks.ResumeLayout(false);
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion
    private System.Windows.Forms.Label lblStatus;
    private MenuStrip menuStrip1;
    private ToolStripMenuItem fileToolStripMenuItem;
    private ToolStripMenuItem importTemplateToolStripMenuItem;
    private ToolStripMenuItem exportTemplateToolStripMenuItem;
    private ToolStripSeparator toolStripSeparator2;
    private ToolStripMenuItem exitToolStripMenuItem;
    private ToolStrip toolStrip1;
    private ToolStripSplitButton toolStripSplitButtonExport;
    private ToolStripMenuItem toolStripMenuItemExportSelected;
    private ToolStripButton toolStripButtonImportExcel;
    private ToolStripButton toolStripMenuItemBlockFilter;
    private ToolStripMenuItem exportAllskip000TotalAmountDueToolStripMenuItem;
    private ToolStripMenuItem exportAllonly000TotalAmountDueToolStripMenuItem;
    private TabPage tabPage2;
    private TabPage tabPage1;
    private TabControl tabBlocks;
    private ToolStripMenuItem settingsToolStripMenuItem;
    private ToolStripMenuItem helpToolStripMenuItem;
    private ToolStripMenuItem checkForUpdatesToolStripMenuItem;
    private ToolStripSeparator toolStripSeparator1;
    private ToolStripMenuItem aboutToolStripMenuItem;
    private ToolStripTextBox toolStripTextBoxSearch;
    private ToolStripButton toolStripButtonClear;
}
