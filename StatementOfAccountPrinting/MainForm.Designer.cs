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
        toolStripSeparator1 = new ToolStripSeparator();
        setToolStripMenuItem = new ToolStripMenuItem();
        toolStripSeparator2 = new ToolStripSeparator();
        aboutToolStripMenuItem = new ToolStripMenuItem();
        toolStripSeparator3 = new ToolStripSeparator();
        exitToolStripMenuItem = new ToolStripMenuItem();
        tabBlocks = new TabControl();
        tabPage1 = new TabPage();
        tabPage2 = new TabPage();
        toolStrip1 = new ToolStrip();
        toolStripSplitButtonExport = new ToolStripSplitButton();
        toolStripMenuItemExportSelected = new ToolStripMenuItem();
        exportAllskip000TotalAmountDueToolStripMenuItem = new ToolStripMenuItem();
        exportAllonly000TotalAmountDueToolStripMenuItem = new ToolStripMenuItem();
        toolStripButtonImportExcel = new ToolStripButton();
        toolStripMenuItemBlockFilter = new ToolStripButton();
        menuStrip1.SuspendLayout();
        tabBlocks.SuspendLayout();
        toolStrip1.SuspendLayout();
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
        menuStrip1.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem });
        menuStrip1.Location = new Point(0, 0);
        menuStrip1.Name = "menuStrip1";
        menuStrip1.Size = new Size(765, 29);
        menuStrip1.TabIndex = 8;
        menuStrip1.Text = "menuStrip1";
        // 
        // fileToolStripMenuItem
        // 
        fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { importTemplateToolStripMenuItem, exportTemplateToolStripMenuItem, toolStripSeparator1, setToolStripMenuItem, toolStripSeparator2, aboutToolStripMenuItem, toolStripSeparator3, exitToolStripMenuItem });
        fileToolStripMenuItem.Name = "fileToolStripMenuItem";
        fileToolStripMenuItem.Size = new Size(46, 25);
        fileToolStripMenuItem.Text = "File";
        // 
        // importTemplateToolStripMenuItem
        // 
        importTemplateToolStripMenuItem.Name = "importTemplateToolStripMenuItem";
        importTemplateToolStripMenuItem.Size = new Size(242, 26);
        importTemplateToolStripMenuItem.Text = "Import Template (.docx)";
        importTemplateToolStripMenuItem.Click += importTemplateToolStripMenuItem_Click;
        // 
        // exportTemplateToolStripMenuItem
        // 
        exportTemplateToolStripMenuItem.Name = "exportTemplateToolStripMenuItem";
        exportTemplateToolStripMenuItem.Size = new Size(242, 26);
        exportTemplateToolStripMenuItem.Text = "Export Template";
        exportTemplateToolStripMenuItem.Click += exportTemplateToolStripMenuItem_Click_1;
        // 
        // toolStripSeparator1
        // 
        toolStripSeparator1.Name = "toolStripSeparator1";
        toolStripSeparator1.Size = new Size(239, 6);
        // 
        // setToolStripMenuItem
        // 
        setToolStripMenuItem.Name = "setToolStripMenuItem";
        setToolStripMenuItem.Size = new Size(242, 26);
        setToolStripMenuItem.Text = "Settings";
        setToolStripMenuItem.Click += settingsToolStripMenuItem_Click;
        // 
        // toolStripSeparator2
        // 
        toolStripSeparator2.Name = "toolStripSeparator2";
        toolStripSeparator2.Size = new Size(239, 6);
        // 
        // aboutToolStripMenuItem
        // 
        aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
        aboutToolStripMenuItem.Size = new Size(242, 26);
        aboutToolStripMenuItem.Text = "About";
        aboutToolStripMenuItem.Click += aboutToolStripMenuItem_Click;
        // 
        // toolStripSeparator3
        // 
        toolStripSeparator3.Name = "toolStripSeparator3";
        toolStripSeparator3.Size = new Size(239, 6);
        // 
        // exitToolStripMenuItem
        // 
        exitToolStripMenuItem.Name = "exitToolStripMenuItem";
        exitToolStripMenuItem.Size = new Size(242, 26);
        exitToolStripMenuItem.Text = "Exit";
        exitToolStripMenuItem.Click += exitToolStripMenuItem_Click;
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
        // toolStrip1
        // 
        toolStrip1.Font = new Font("Segoe UI", 12F);
        toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripSplitButtonExport, toolStripButtonImportExcel, toolStripMenuItemBlockFilter });
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
        toolStrip1.ResumeLayout(false);
        toolStrip1.PerformLayout();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion
    private System.Windows.Forms.Label lblStatus;
    private MenuStrip menuStrip1;
    private ToolStripMenuItem fileToolStripMenuItem;
    private ToolStripMenuItem importTemplateToolStripMenuItem;
    private ToolStripMenuItem exportTemplateToolStripMenuItem;
    private TabControl tabBlocks;
    private TabPage tabPage1;
    private TabPage tabPage2;
    private ToolStripSeparator toolStripSeparator1;
    private ToolStripMenuItem setToolStripMenuItem;
    private ToolStripSeparator toolStripSeparator2;
    private ToolStripMenuItem exitToolStripMenuItem;
    private ToolStrip toolStrip1;
    private ToolStripSplitButton toolStripSplitButtonExport;
    private ToolStripMenuItem toolStripMenuItemExportSelected;
    private ToolStripButton toolStripButtonImportExcel;
    private ToolStripButton toolStripMenuItemBlockFilter;
    private ToolStripMenuItem aboutToolStripMenuItem;
    private ToolStripSeparator toolStripSeparator3;
    private ToolStripMenuItem exportAllskip000TotalAmountDueToolStripMenuItem;
    private ToolStripMenuItem exportAllonly000TotalAmountDueToolStripMenuItem;
}
