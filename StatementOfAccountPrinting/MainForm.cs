using ClosedXML.Excel;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.ComponentModel;

namespace StatementOfAccountPrinting;

public partial class MainForm : Form
{
    private readonly BindingList<CemeteryLotAccount> _accounts = new();
    private const string AppTitle = "Statement of Account Printing — Public Cemetery Lot Rental";
    private const string TemplateFileName = "WordTemplate.docx";
    private const string StatementPrefix = "SOA_";
    private readonly List<Stream> _fontStreams = new();
    private BlockFilterPopup? _blockPopup;
    private readonly HashSet<string> _selectedBlocks = new(StringComparer.OrdinalIgnoreCase);
    private bool _blockFilterInitialized;   // set true once the user has made a selection

    public MainForm()
    {
        InitializeComponent();

        FormClosing += MainForm_FormClosing;
        btnPrintSelected.Click += btnPrintSelected_Click;

        var wa = Screen.FromControl(this).WorkingArea;

        int newWidth = Math.Min(Width, wa.Width);
        int newHeight = Math.Min(Height, wa.Height);

        int newLeft = Math.Max(wa.Left, Math.Min(Left, wa.Right - newWidth));
        int newTop = Math.Max(wa.Top, Math.Min(Top, wa.Bottom - newHeight));

        Bounds = new Rectangle(newLeft, newTop, newWidth, newHeight);

        Text = AppTitle;
        //dgvData.DataSource = _accounts;
        //RebuildBlockTabs();
        QuestPDF.Settings.License = LicenseType.Community;
        RegisterFonts();

        ToolTip toolTip = new ToolTip();
        toolTip.SetToolTip(btnBlockFilter, "Select Blocks to Print");
        toolTip.SetToolTip(btnPrintPdf, "Export as PDF");
        toolTip.SetToolTip(btnPrintSelected, "Print selected row(s) from the active tab");

        var savedFormat = Properties.Settings.Default.ExportFormat;
        btnPrintPdf.Text = $"Export as {savedFormat}";
        bool isPdf = savedFormat == "PDF";
        fileToolStripMenuItem.Visible = !isPdf;
    }

    private async void MainForm_Shown(object sender, EventArgs e)
    {
        await LoadAccountsFromDiskAsync();
    }
    private void RegisterFonts()
    {
        // ── Calibri (existing) ──
        foreach (var path in new[]
        {
        "Fonts/calibri.ttf",
        "Fonts/calibrib.ttf",
        "Fonts/calibrii.ttf",
    })
        {
            var stream = File.OpenRead(path);
            _fontStreams.Add(stream);
            FontManager.RegisterFontFromStream(stream);
        }

        // ── Times New Roman (system font — register from Windows\Fonts) ──
        string winFonts = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        foreach (var file in new[] { "times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf" })
        {
            var full = Path.Combine(winFonts, file);
            if (!File.Exists(full)) continue;

            var stream = File.OpenRead(full);
            _fontStreams.Add(stream);
            FontManager.RegisterFontFromStream(stream);
        }
    }
    private void RebuildBlockTabs()
    {
        // Preserve the currently selected block, if any
        string? current = tabBlocks.SelectedTab?.Tag as string;

        tabBlocks.TabPages.Clear();

        if (_accounts.Count == 0) return;

        // ── Determine which blocks should be visible ──
        var allBlockNames = _accounts
            .Select(a => string.IsNullOrWhiteSpace(a.LotBlockNumber)
                            ? "(No Block)"
                            : a.LotBlockNumber.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        // If nothing is selected, or everything is selected → show all
        bool showAll = _selectedBlocks.Count == 0
               || _selectedBlocks.Count >= allBlockNames.Count;

        var visibleBlocks = showAll
            ? allBlockNames
            : allBlockNames.Where(b => _selectedBlocks.Contains(b)).ToList();

        // ── Group only the visible blocks ──
        var groups = _accounts
            .Where(a =>
            {
                var key = string.IsNullOrWhiteSpace(a.LotBlockNumber)
                            ? "(No Block)"
                            : a.LotBlockNumber.Trim();
                return visibleBlocks.Contains(key, StringComparer.OrdinalIgnoreCase);
            })
            .GroupBy(a => string.IsNullOrWhiteSpace(a.LotBlockNumber)
                            ? "(No Block)"
                            : a.LotBlockNumber.Trim())
            .OrderBy(g => ExtractBlockNumber(g.Key))          // 1, 2, 10, 100, …
            .ThenBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (var group in groups)
        {
            var page = new TabPage
            {
                Text = $"{group.Key} ({group.Count()})",
                Tag = group.Key,
                Padding = new Padding(4)
            };

            // A binding list scoped to this block, so the grid updates as items change
            var blockList = new BindingList<CemeteryLotAccount>(group.ToList());

            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = true,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,           // ← multi-select support
                DataSource = blockList
            };

            grid.SelectionChanged += (s, e) =>
            {
                btnPrintSelected.Enabled = grid.SelectedRows.Count > 0;
            };

            // Wire the context menu
            var menu = new ContextMenuStrip();
            menu.Items.Add("Print Selected Record(s)", null,
                (s, e) => btnPrintSelected_Click(s, e));
            grid.ContextMenuStrip = menu;

            page.Controls.Add(grid);
            tabBlocks.TabPages.Add(page);
        }

        // ── Subscribe ONCE for the tab-change event (not inside the loop) ──
        tabBlocks.SelectedIndexChanged -= TabBlocks_SelectedIndexChanged;
        tabBlocks.SelectedIndexChanged += TabBlocks_SelectedIndexChanged;

        // Restore previous selection if still present
        if (current != null)
        {
            foreach (TabPage p in tabBlocks.TabPages)
            {
                if (string.Equals(p.Tag as string, current, StringComparison.OrdinalIgnoreCase))
                {
                    tabBlocks.SelectedTab = p;
                    break;
                }
            }
        }

        // Fallback: if the previous tab is gone, select the first available tab
        if (tabBlocks.TabPages.Count > 0 && tabBlocks.SelectedTab == null)
            tabBlocks.SelectedTab = tabBlocks.TabPages[0];

        // If there are no tabs at all (e.g., a filter matched nothing), disable print
        if (tabBlocks.TabPages.Count == 0)
            btnPrintSelected.Enabled = false;
    }

    private void TabBlocks_SelectedIndexChanged(object? sender, EventArgs e)
    {
        var g = GetActiveGrid();
        btnPrintSelected.Enabled = g?.SelectedRows.Count > 0;
    }
    private DataGridView? GetActiveGrid()
    {
        if (tabBlocks.SelectedTab == null) return null;
        return tabBlocks.SelectedTab.Controls.OfType<DataGridView>().FirstOrDefault();
    }
    private List<CemeteryLotAccount> GetSelectedAccountsFromActiveGrid()
    {
        var grid = GetActiveGrid();
        if (grid == null) return new List<CemeteryLotAccount>();

        // If nothing is selected, return an empty list (caller can decide to fall back to all)
        if (grid.SelectedRows.Count == 0) return new List<CemeteryLotAccount>();

        return grid.SelectedRows
            .Cast<DataGridViewRow>()
            .Select(r => r.DataBoundItem as CemeteryLotAccount)
            .Where(a => a != null)
            .Cast<CemeteryLotAccount>()
            .ToList();
    }
    private async void btnPrintSelected_Click(object sender, EventArgs e)
    {
        btnImportExcel.Enabled = false;
        btnGenerateWord.Enabled = false;
        btnPrintPdf.Enabled = false;
        btnPrintSelected.Enabled = false;

        try
        {
            // ── Collect selected rows (or all rows from the active tab) ──
            var selected = GetSelectedAccountsFromActiveGrid();

            if (selected.Count == 0)
            {
                var grid = GetActiveGrid();
                if (grid?.DataSource is IEnumerable<CemeteryLotAccount> all)
                    selected = all.ToList();
            }

            if (selected.Count == 0)
            {
                MessageBox.Show("No records to print.",
                    "Print", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // ── Read settings once ──
            string paperSize = Properties.Settings.Default.PaperSize ?? "A4";
            bool skipZero = Properties.Settings.Default.SkipZeroTotal;
            string exportFormat = Properties.Settings.Default.ExportFormat ?? "PDF";

            // Apply skip-zero filter if enabled
            if (skipZero)
                selected = selected.Where(a => !IsZeroTotal(a.TotalAmountDue)).ToList();

            if (selected.Count == 0)
            {
                MessageBox.Show("All selected records have ₱0.00 total.",
                    "Print", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // ── Ask for an output folder ──
            using var fbd = new FolderBrowserDialog
            {
                Description = "Choose output folder for selected record(s)"
            };
            if (fbd.ShowDialog() != DialogResult.OK) return;

            var outputFolder = fbd.SelectedPath;

            // ── File name base ──
            string fileBase = selected
                .Select(a => a.LotBlockNumber?.Trim() ?? "(No Block)")
                .Distinct()
                .Count() == 1
                    ? selected[0].LotBlockNumber!.Trim()
                    : "Selected";

            string safeBase = MakeSafeFileName(fileBase);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            bool isWord = string.Equals(exportFormat, "DOCX", StringComparison.OrdinalIgnoreCase);
            string extension = isWord ? "docx" : "pdf";
            string fileName = $"Statements_{safeBase}_{paperSize}_{timestamp}.{extension}";
            string filePath = Path.Combine(outputFolder, fileName);

            // ── Run the export with progress ──
            var sw = System.Diagnostics.Stopwatch.StartNew();

            await ProgressDialog.RunAsync(this,
                isWord ? "Generating DOCX…" : "Generating PDF…",
                async progress =>
                {
                    await Task.Run(() =>
                    {
                        progress.Report((10, $"Preparing {selected.Count} record(s)…"));

                        if (isWord)
                        {
                            // Word: one combined .docx for the selected rows
                            WordTemplateService.GenerateCombinedWord(filePath, selected.ToList(), paperSize);
                        }
                        else
                        {
                            // PDF: reuse the per-block renderer
                            CreatePdfPerBlock(filePath, paperSize, skipZero, selected);
                        }

                        progress.Report((100, $"Done. Saved {selected.Count} record(s)."));
                    });
                }, keepOpenWhenDone: false);

            sw.Stop();
            lblStatus.Text =
                $"Saved {selected.Count} record(s) as {exportFormat} in {FormatElapsed(sw.Elapsed)} → {filePath}";

            if (AskYesNo($"{exportFormat} created. Open it now?"))
                OpenFile(filePath);
        }
        catch (Exception ex)
        {
            ShowError("Print failed", ex);
        }
        finally
        {
            btnImportExcel.Enabled = true;
            btnGenerateWord.Enabled = _accounts.Count > 0;
            btnPrintPdf.Enabled = _accounts.Count > 0;
            btnPrintSelected.Enabled = _accounts.Count > 0;
        }
    }
    private static int ExtractBlockNumber(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return int.MaxValue;

        // Find the first run of digits anywhere in the string
        var match = System.Text.RegularExpressions.Regex.Match(name, @"\d+");
        if (match.Success && int.TryParse(match.Value, out var n))
            return n;

        return int.MaxValue;
    }
    private void PrependAllTab()
    {
        if (_accounts.Count == 0) return;

        var allPage = new TabPage
        {
            Text = $"ALL ({_accounts.Count})",
            Tag = null
        };

        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = true,
            DataSource = _accounts
        };

        allPage.Controls.Add(grid);
        tabBlocks.TabPages.Insert(0, allPage);
    }

    #region IMPORT...
    private async void btnImportExcel_Click(object sender, EventArgs e)
    {
        btnImportExcel.Enabled = false;
        btnGenerateWord.Enabled = false;
        btnPrintPdf.Enabled = false;
        btnPrintSelected.Enabled = false;   // ← add this
        btnBlockFilter.Enabled = false;

        using var ofd = new OpenFileDialog
        {
            Title = "Select Excel file",
            Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        };
        if (ofd.ShowDialog() != DialogResult.OK) return;

        var filePath = ofd.FileName;
        var imported = new List<CemeteryLotAccount>();   // ← LOCAL list, not _accounts

        btnImportExcel.Enabled = false;

        try
        {
            await ProgressDialog.RunAsync(this, "Importing Excel…", async progress =>
            {
                await Task.Run(() =>
                {
                    progress.Report((0, "Opening workbook…"));

                    using var wb = new XLWorkbook(filePath);

                    foreach (var ws in wb.Worksheets)
                    {
                        var sheetname = GetSheetName(ws);
                        if (!sheetname.StartsWith("BLOCK") ||
                            !ws.Name.TrimStart().StartsWith("BLOCK", StringComparison.OrdinalIgnoreCase))
                            continue;

                        var used = ws.RangeUsed();
                        if (used == null) continue;

                        int firstRow = used.FirstRow().RowNumber();
                        int lastRow = used.LastRow().RowNumber();

                        for (int rowNum = firstRow; rowNum <= lastRow; rowNum++)
                        {
                            var row = ws.Row(rowNum);

                            var idxCell = row.Cell("A");
                            if (idxCell.IsEmpty()) continue;
                            if (!idxCell.Value.IsNumber) continue;

                            var lessee = row.Cell("B").GetString().Trim();
                            if (string.IsNullOrWhiteSpace(lessee)) continue;

                            decimal rental = ParseNum(row, "V") + ParseNum(row, "W") + ParseNum(row, "X");
                            decimal surcharge = ParseNum(row, "Z") + ParseNum(row, "AA") + ParseNum(row, "AB");
                            decimal interest = ParseNum(row, "AD") + ParseNum(row, "AE") + ParseNum(row, "AF");
                            decimal total = rental + surcharge + interest;

                            // ⬇ Add to the LOCAL list, not _accounts
                            imported.Add(new CemeteryLotAccount
                            {
                                DateGenerated = DateTime.Today.ToString("MMMM dd, yyyy"),
                                LesseeName = lessee,
                                NameOfDeceased = row.Cell("C").GetString().Trim(),
                                Address = row.Cell("D").GetString().Trim(),
                                LotBlockNumber = sheetname,
                                AreaSqm = row.Cell("E").GetString().Trim(),
                                RentalOrRenewalFee = rental.ToString("0.00"),
                                Surcharge = surcharge.ToString("0.00"),
                                Interest = interest.ToString("0.00"),
                                TotalAmountDue = total.ToString("0.00"),
                            });
                        }
                    }

                    progress.Report((100, "Finishing up…"));
                });
            }, keepOpenWhenDone: false);

            // ── Back on the UI thread — now safe to touch _accounts and dgvData ──
            _accounts.Clear();
            foreach (var a in imported)
                _accounts.Add(a);

            //PopulateBlockFilter();
            RebuildBlockTabs();
            UpdateBlockFilterCaption();
            SaveAccountsToDisk();

            btnGenerateWord.Enabled = btnPrintPdf.Enabled = _accounts.Count > 0;
            btnPrintSelected.Enabled = _accounts.Count > 0;   // ← add this
            lblStatus.Text = $"Imported {_accounts.Count} row(s) from " +
                             $"{Path.GetFileName(filePath)}.";

            if (btnPrintPdf.Enabled && btnPrintPdf.Visible)
            {
                btnBlockFilter.Enabled = true;
                btnPrintPdf.Focus();
                btnPrintPdf.Select();
            }
        }
        catch (Exception ex)
        {
            ShowError("Import failed", ex);
        }
        finally
        {
            btnImportExcel.Enabled = true;
            btnGenerateWord.Enabled = _accounts.Count > 0;
            btnPrintPdf.Enabled = _accounts.Count > 0;
            btnPrintSelected.Enabled = _accounts.Count > 0;
            btnBlockFilter.Enabled = _accounts.Count > 0;
        }
    }
    #endregion
    private static decimal ParseNum(IXLRow row, string columnLetter)
    {
        try
        {
            var cell = row.Cell(columnLetter);
            if (cell == null || cell.IsEmpty()) return 0m;

            // Try cached numeric value first
            if (cell.HasFormula)
            {
                var cached = cell.CachedValue;
                if (!cached.IsBlank && cached.Type == XLDataType.Number)
                    return (decimal)cached.GetNumber();
            }

            if (!cell.Value.IsBlank && cell.Value.Type == XLDataType.Number)
                return (decimal)cell.Value.GetNumber();

            // Fall back to parsing the display string
            var s = cell.GetFormattedString()?.Trim() ?? "";
            s = s.Replace("₱", "").Replace(",", "").Trim();
            return decimal.TryParse(s, out var d) ? d : 0m;
        }
        catch
        {
            return 0m;
        }
    }
    private static string GetCellValue(IXLRow row, string columnLetter)
    {
        try
        {
            var cell = row.Cell(columnLetter);
            if (cell == null || cell.IsEmpty()) return string.Empty;

            // Prefer cached formula result when available
            if (cell.HasFormula)
            {
                var cached = cell.CachedValue;
                if (!cached.IsBlank)
                    return FormatCellValue(cached);
            }

            if (!cell.Value.IsBlank)
                return FormatCellValue(cell.Value);

            return cell.GetFormattedString()?.Trim() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }
    private static string FormatCellValue(XLCellValue v)
    {
        try
        {
            return v.Type switch
            {
                XLDataType.Number => v.GetNumber().ToString("0.00"),
                XLDataType.DateTime => v.GetDateTime().ToString("yyyy-MM-dd"),
                XLDataType.Boolean => v.GetBoolean().ToString(),
                XLDataType.Text => v.GetText()?.Trim() ?? string.Empty,
                XLDataType.Blank => string.Empty,
                XLDataType.Error => string.Empty,
                _ => v.ToString() ?? string.Empty,
            };
        }
        catch
        {
            return string.Empty;
        }
    }
    private static string GetCell(IXLRangeRow row, int col)
    {
        try
        {
            var cell = row.Cell(col);
            if (cell == null || cell.IsEmpty()) return string.Empty;

            // Preserve what the user sees (not the underlying number)
            return cell.GetFormattedString()?.Trim() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string GetCellByLetter(IXLRangeRow row, string columnLetter)
    {
        try
        {
            var cell = row.Cell(columnLetter);   // "U", "Y", "AC", "AH"
            if (cell == null || cell.IsEmpty()) return string.Empty;
            return cell.GetFormattedString()?.Trim() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string Norm(string s)
    {
        if (string.IsNullOrEmpty(s)) return string.Empty;
        var sb = new System.Text.StringBuilder(s.Length);
        foreach (var ch in s)
        {
            if (char.IsWhiteSpace(ch) || ch == '\u00A0') continue;
            sb.Append(char.ToUpperInvariant(ch));
        }
        return sb.ToString();
    }

    // Get a meaningful sheet name (falls back to A2 title if tab is generic)
    private static string GetSheetName(IXLWorksheet ws)
    {
        var tab = ws.Name?.Trim() ?? string.Empty;
        //if (!string.IsNullOrWhiteSpace(tab) &&
        //    !tab.StartsWith("Sheet", StringComparison.OrdinalIgnoreCase) &&
        //    !int.TryParse(tab, out _))
        //{
        //    return tab;
        //}

        var a2 = ws.Cell("A2").GetString();
        // Strip the "INVENTORY OF CEMETERY..." prefix
        var cleaned = a2
            .Replace("INVENTORY OF CEMETERY", "", StringComparison.OrdinalIgnoreCase)
            .Replace("LOT", "", StringComparison.OrdinalIgnoreCase)
            .Replace("\t", "").Replace("\r", "").Replace("\n", "")
            .Trim(' ', '-', ':')
            .Trim();

        return string.IsNullOrWhiteSpace(cleaned) ? tab : cleaned;
    }

    // ──────────── 2. Create starter Word template ────────────
    private void btnCreateTemplate_Click(object sender, EventArgs e)
    {
        ExportTemplate();
    }
    private void ExportTemplate()
    {
        using var sfd = new SaveFileDialog
        {
            Title = "Save blank template as",
            FileName = WordTemplateService.TemplateFileName,
            DefaultExt = "docx",
            Filter = "Word documents (*.docx)|*.docx|All files (*.*)|*.*",
            AddExtension = true,
            OverwritePrompt = true
        };

        if (sfd.ShowDialog() != DialogResult.OK) return;

        exportTemplateToolStripMenuItem.Enabled = false;
        Cursor = Cursors.WaitCursor;
        try
        {
            WordTemplateService.CreateTemplate(sfd.FileName);

            lblStatus.Text = $"Blank template saved: {Path.GetFileName(sfd.FileName)}";

            if (AskYesNo(
                "Blank template saved.\n\n" +
                "You can now edit it in Word. If you want the app to use your edited " +
                "version next time, replace this file:\n\n" +
                $"  {WordTemplateService.ResolveTemplatePath()}"))
            {
                OpenFile(sfd.FileName);
            }
        }
        catch (IOException ioEx)
        {
            MessageBox.Show(
                "The destination file is open in another program (likely Word).\n" +
                "Close it and try again.\n\n" +
                $"Details: {ioEx.Message}",
                "Template creation failed",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
        catch (Exception ex)
        {
            ShowError("Template creation failed", ex);
        }
        finally
        {
            Cursor = Cursors.Default;
            exportTemplateToolStripMenuItem.Enabled = true;
        }
    }

    // ───── 3. Fill template → one .docx statement per row ─────
    private void btnGenerateWord_Click(object sender, EventArgs e)
    {
        ExportAsWord();
    }
    private async Task ExportAsWord()
    {
        using var sfd = new SaveFileDialog
        {
            Title = "Save statements as",
            FileName = $"Statements_{DateTime.Now:yyyyMMdd_HHmmss}.docx",
            DefaultExt = "docx",
            Filter = "Word documents (*.docx)|*.docx",
            AddExtension = true,
            OverwritePrompt = true
        };
        if (sfd.ShowDialog() != DialogResult.OK) return;

        if (_accounts.Count == 0)
        {
            MessageBox.Show("No accounts loaded.", "Generate Word",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        try
        {
            string paperSize = Properties.Settings.Default.PaperSize ?? "A4";
            WordTemplateService.GenerateCombinedWord(sfd.FileName, _accounts.ToList(), paperSize);

            lblStatus.Text = $"Generated {_accounts.Count} statement(s) in {Path.GetFileName(sfd.FileName)}";

            var templatePath = WordTemplateService.ResolveTemplatePath();
            var headerLines = SoaTemplateReader.ReadHeaders(templatePath);
            var footerLines = SoaTemplateReader.ReadFooters(templatePath);

            if (AskYesNo("Word file created. Open it now?"))
                OpenFile(sfd.FileName);
        }
        catch (Exception ex)
        {
            ShowError("Word generation failed", ex);
        }
    }
    private async void btnPrintPdf_Click(object sender, EventArgs e)
    {
        string format = Properties.Settings.Default.ExportFormat ?? "PDF";

        if (string.Equals(format, "DOCX", StringComparison.OrdinalIgnoreCase))
        {
            ExportAsWord();
        }
        else
        {
            await ExportAsPdfAsync();
        }
    }

    private async Task ExportAsPdfAsync()
    {
        btnImportExcel.Enabled = false;
        btnGenerateWord.Enabled = false;
        btnPrintPdf.Enabled = false;
        btnPrintSelected.Enabled = false;   // ← add this
        btnBlockFilter.Enabled = false;

        if (_accounts.Count == 0)
        {
            MessageBox.Show("No accounts to export.", "PDF",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        string paperSize = Properties.Settings.Default.PaperSize ?? "A4";
        bool skipZero = Properties.Settings.Default.SkipZeroTotal;

        using var fbd = new FolderBrowserDialog
        {
            Description = "Choose output folder for per-block PDFs"
        };
        if (fbd.ShowDialog() != DialogResult.OK) return;

        var outputFolder = fbd.SelectedPath;

        // ── Read selected blocks from the popup's stored set ──
        var selectedBlocks = new HashSet<string>(_selectedBlocks, StringComparer.OrdinalIgnoreCase);

        // ── Build the filtered source as a List<CemeteryLotAccount> ──
        List<CemeteryLotAccount> source = _accounts.ToList();

        // 1) Block filter
        if (selectedBlocks.Count > 0)
        {
            int totalBlocksInForm = source
                .Select(a => string.IsNullOrWhiteSpace(a.LotBlockNumber)
                                ? "(No Block)"
                                : a.LotBlockNumber.Trim())
                .Distinct()
                .Count();

            // Only apply the filter if it actually narrows things down
            if (selectedBlocks.Count < totalBlocksInForm)
            {
                source = source
                    .Where(a =>
                    {
                        var key = string.IsNullOrWhiteSpace(a.LotBlockNumber)
                                    ? "(No Block)"
                                    : a.LotBlockNumber.Trim();
                        return selectedBlocks.Contains(key);
                    })
                    .ToList();
            }
        }

        var skippedZeroOrigCount = source.Count();

        // 2) Skip-zero filter
        if (skipZero)
            source = source.Where(a => !IsZeroTotal(a.TotalAmountDue)).ToList();

        if (source.Count == 0)
        {
            MessageBox.Show("No records match the current filters.",
                "PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // 3) Group by block — this is where `groups` comes from
        var groups = source
            .GroupBy(a => string.IsNullOrWhiteSpace(a.LotBlockNumber)
                            ? "(No Block)"
                            : a.LotBlockNumber.Trim())
            .OrderBy(g => ExtractBlockNumber(g.Key))
            .ThenBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();

        int totalBlocks = groups.Count;
        int totalRecords = source.Count;

        btnImportExcel.Enabled = btnGenerateWord.Enabled = btnPrintPdf.Enabled = false;

        var sw = System.Diagnostics.Stopwatch.StartNew();

        try
        {
            await ProgressDialog.RunAsync(this, "Generating PDFs…", async progress =>
            {
                var cts = new CancellationTokenSource();
                int estimatedMs = Math.Max(1500, totalBlocks * 800);

                var ticker = Task.Run(async () =>
                {
                    var inner = System.Diagnostics.Stopwatch.StartNew();
                    int lastPct = 0;

                    while (!cts.IsCancellationRequested && inner.ElapsedMilliseconds < estimatedMs)
                    {
                        double ratio = inner.ElapsedMilliseconds / (double)estimatedMs;
                        int pct = Math.Min(95, (int)(ratio * 95));

                        if (pct > lastPct)
                        {
                            lastPct = pct;
                            int done = Math.Min(totalBlocks, (int)Math.Ceiling(ratio * totalBlocks));
                            progress.Report((pct,
                                $"Generating {done} of {totalBlocks} block PDFs…  ({FormatElapsed(sw.Elapsed)} elapsed)"));
                        }

                        try { await Task.Delay(50, cts.Token); }
                        catch (TaskCanceledException) { break; }
                    }

                    if (!cts.IsCancellationRequested)
                        progress.Report((95,
                            $"Finalizing {totalBlocks} file(s)…  ({FormatElapsed(sw.Elapsed)} elapsed)"));
                });

                int generated = 0;
                try
                {
                    await Task.Run(() =>
                    {
                        foreach (var group in groups)
                        {
                            string safeBlock = MakeSafeFileName(group.Key);
                            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                            string fileName = $"Statements_{safeBlock}_{paperSize}_{timestamp}.pdf";
                            string blockPath = Path.Combine(outputFolder, fileName);

                            // Same template, only that block's rows
                            CreatePdfPerBlock(blockPath, paperSize, skipZero, group.ToList());

                            generated++;
                        }
                    });
                }
                finally
                {
                    cts.Cancel();
                    try { await ticker; } catch { }
                }

                progress.Report((100,
                    $"PDF export complete.\n" +
                    $"Folder      : {outputFolder}\n" +
                    $"Blocks      : {generated} of {totalBlocks}\n" +
                    $"Records     : {totalRecords} of {skippedZeroOrigCount}" +
                        (skipZero ? $" (skipped ₱0.00)" : "") + "\n" +
                    $"Paper size  : {paperSize}\n" +
                    $"Elapsed time: {FormatElapsed(sw.Elapsed)}"));
            });

            sw.Stop();
            lblStatus.Text = $"Saved {totalBlocks} block PDF(s) in {FormatElapsed(sw.Elapsed)} → {outputFolder}";

            if (AskYesNo("Open the output folder?"))
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                { FileName = outputFolder, UseShellExecute = true });

            // Reset block filter after export
            _selectedBlocks.Clear();
            _blockFilterInitialized = false;
            UpdateBlockFilterCaption();
        }
        catch (Exception ex)
        {
            sw.Stop();
            ShowError($"PDF creation failed after {FormatElapsed(sw.Elapsed)}", ex);
        }
        finally
        {
            btnImportExcel.Enabled = true;
            btnGenerateWord.Enabled = _accounts.Count > 0;
            btnPrintPdf.Enabled = _accounts.Count > 0;
            btnPrintSelected.Enabled = _accounts.Count > 0;
            btnBlockFilter.Enabled = _accounts.Count > 0;
        }
    }

    private static string FormatElapsed(TimeSpan t)
    {
        if (t.TotalHours >= 1) return $"{(int)t.TotalHours}h {t.Minutes}m {t.Seconds}s";
        if (t.TotalMinutes >= 1) return $"{t.Minutes}m {t.Seconds}s";
        if (t.TotalSeconds >= 1) return $"{t.TotalSeconds:0.0}s";
        return $"{t.TotalMilliseconds:0}ms";
    }

    private void CreatePdfPerBlock(string path, string paperSize, bool skipZero, IEnumerable<CemeteryLotAccount> accounts)
    {
        var list = accounts.ToList();
        if (list.Count == 0)
        {
            MessageBox.Show("No accounts to export.", "PDF",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // Resolve logos
        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        var logoLeft = Path.Combine(baseDir, "Templates", "Images", "navotas-logo-left.jpg");
        var logoRight = Path.Combine(baseDir, "Templates", "Images", "blgf-logo-right.png");

        byte[]? logoLeftBytes = File.Exists(logoLeft) ? File.ReadAllBytes(logoLeft) : null;
        byte[]? logoRightBytes = File.Exists(logoRight) ? File.ReadAllBytes(logoRight) : null;

        Document.Create(doc =>
        {
            foreach (var a in list)          // ← only the accounts passed in
            {
                doc.Page(page =>
                {
                    // ── Same layout as your original CreatePdf ──
                    switch (paperSize)
                    {
                        case "Legal": page.Size(PageSizes.Legal); break;
                        default: page.Size(PageSizes.A4); break;
                    }

                    page.MarginTop(1, Unit.Centimetre);
                    page.MarginLeft(2.54f, Unit.Centimetre);
                    page.MarginRight(2.54f, Unit.Centimetre);
                    page.MarginBottom(0.5f, Unit.Centimetre);

                    page.DefaultTextStyle(t => t
                        .FontFamily("Lato")
                        .FontSize(9)
                        .FontColor(Colors.Black));

                    // ── HEADER ──
                    page.Header().PaddingTop(0).Column(h =>
                    {
                        h.Item().Row(row =>
                        {
                            row.ConstantItem(65).Height(65).AlignLeft().AlignMiddle().Element(e =>
                            {
                                if (logoLeftBytes != null) e.Image(logoLeftBytes).FitArea();
                            });

                            row.RelativeItem().AlignCenter().AlignMiddle().Column(center =>
                            {
                                center.Item().AlignCenter()
                                    .Text("Republic of the Philippines")
                                    .FontSize(11).FontFamily("Times New Roman");

                                center.Item().AlignCenter()
                                    .Text("CITY OF NAVOTAS")
                                    .FontSize(16).Bold().FontFamily("Times New Roman")
                                    .FontColor(Colors.Grey.Darken3);

                                center.Item().AlignCenter()
                                    .Text("OFFICE OF THE CITY TREASURER")
                                    .FontSize(12).FontFamily("Times New Roman");
                            });

                            row.ConstantItem(65).Height(65).AlignRight().AlignMiddle().Element(e =>
                            {
                                if (logoRightBytes != null) e.Image(logoRightBytes).FitArea();
                            });
                        });

                        h.Item().PaddingTop(20).AlignCenter()
                            .Text("STATEMENT OF ACCOUNT – PUBLIC CEMETERY LOT RENTAL")
                            .FontSize(9).Bold();

                        h.Item().PaddingTop(2).AlignCenter()
                            .Text("(Pahayag ng Bayarin sa Renta ng Lupa sa Navotas City Public Cemetery)")
                            .FontSize(9).Bold();

                        h.Item().PaddingTop(14);
                    });

                    // ── CONTENT (same as your existing CreatePdf) ──
                    page.Content().Column(col =>
                    {
                        col.Spacing(0);

                        // Info block
                        col.Item().Row(row =>
                        {
                            row.RelativeItem().Column(inf =>
                            {
                                InfoLine(inf, "Date / Petsa: ", a.DateGenerated);
                                inf.Item().PaddingTop(12);
                                InfoLine(inf, "Name of Lessee / Pangalan: ", a.LesseeName);
                                InfoLine(inf, "Address / Tirahan: ", a.Address);
                                InfoLine(inf, "Block / Lot No.: ", a.LotBlockNumber);
                                InfoLine(inf, "Area / Sukat: ", a.AreaSqm + " sq. m.");
                                InfoLine(inf, "Name of Deceased: ", a.NameOfDeceased);
                            });
                        });

                        // Intro paragraph
                        col.Item().PaddingTop(12).Text(txt =>
                        {
                            txt.Span("Batay sa talaan ng ").FontSize(9);
                            txt.Span("City Treasurer’s Office, ").FontSize(9);
                            txt.Span("ang sumusunod na bayarin ay ").FontSize(9);
                            txt.Span("nakatakda ").FontSize(9);
                            txt.Span("para sa ").FontSize(9);
                            txt.Span("renta o renewal ").FontSize(9);
                            txt.Span("ng inyong lote sa ").FontSize(9);
                            txt.Span("Navotas City Public Cemetery:").FontSize(9).Bold();
                        });

                        // Charges table
                        col.Item()
                           .PaddingTop(12)
                           .AlignCenter()
                           .Width(300)
                           .Table(table =>
                           {
                               table.ColumnsDefinition(c =>
                               {
                                   c.RelativeColumn(3);
                                   c.RelativeColumn(2);
                               });

                               table.Cell().Border(0.5f).Padding(4)
                                    .AlignCenter().AlignMiddle()
                                    .Text("PARTICULARS / DETALYE").FontSize(9).Bold();

                               table.Cell().Border(0.5f).Padding(4)
                                    .AlignCenter().AlignMiddle()
                                    .Text("AMOUNT / HALAGA").FontSize(9).Bold();

                               ChargeRow(table, "Rental / Renewal Fee", a.RentalOrRenewalFee);
                               ChargeRow(table, "Surcharge, if applicable", a.Surcharge);
                               ChargeRow(table, "Interest, if applicable", a.Interest);

                               table.Cell().Border(0.5f).Padding(4).AlignLeft().AlignMiddle()
                                    .Element(cell =>
                                    {
                                        cell.Text(txt =>
                                        {
                                            txt.Span("Total Amount Due / ").FontSize(9);
                                            txt.Span("Kabuuang Babayaran").FontSize(9).Italic();
                                        });
                                    });

                               table.Cell().Border(0.5f).Padding(4).AlignRight()
                                    .Text("₱ " + a.TotalAmountDue).FontSize(9);
                           });

                        // Instructions (same as before)
                        col.Item().PaddingTop(24).AlignCenter()
                            .Text("PAYMENT AND INQUIRY INSTRUCTIONS").FontSize(9).Bold();

                        col.Item().PaddingTop(2).AlignCenter()
                            .Text("(Paraan ng Pagbabayad at Pagtatanong)").FontSize(9).Bold();

                        col.Item().PaddingTop(12).Text(txt =>
                        {
                            txt.Span("For Payment: ").FontSize(9).Bold();
                            txt.Span("Present this Statement of Account and ").FontSize(9);
                            txt.Span("pay immediately ").FontSize(9).Bold();
                            txt.Span("the amount due at the ").FontSize(9);
                            txt.Span("authorized collection counter of the City Treasurer’s Office. ").FontSize(9).Bold();
                            txt.Span("An Official Receipt shall be issued upon payment.").FontSize(9);
                        });

                        col.Item().PaddingTop(6).Text(txt =>
                        {
                            txt.Span("Para sa Pagbabayad: ").FontSize(9).Bold();
                            txt.Span("Dalhin ang Statement of Account na ito at agad na bayaran ang kaukulang halaga sa ").FontSize(9);
                            txt.Span("awtorisadong collection counter ng City Treasurer’s Office. ").FontSize(9).Bold();
                            txt.Span("Humingi at ingatan ang Official Receipt bilang patunay ng pagbabayad.").FontSize(9);
                        });

                        col.Item().PaddingTop(6).Text(txt =>
                        {
                            txt.Span("For Inquiries / Record Verification: ").FontSize(9).Bold();
                            txt.Span("Please coordinate with the ").FontSize(9);
                            txt.Span("City Treasurer’s Office or Navotas City Public Cemetery ").FontSize(9).Bold();
                            txt.Span("regarding your account, payment history, lease period, or renewal.").FontSize(9);
                        });

                        col.Item().PaddingTop(4).Text(txt =>
                        {
                            txt.Span("Para sa Tanong / Beripikasyon: ").FontSize(9).Bold();
                            txt.Span("Makipag-ugnayan sa ").FontSize(9);
                            txt.Span("City Treasurer’s Office or Navotas City Public Cemetery ").FontSize(9).Bold();
                            txt.Span("tungkol sa inyong account, payment record, lease period, o renewal.").FontSize(9);
                        });

                        col.Item().PaddingTop(6).Text(txt =>
                        {
                            txt.Span("Paalala: ").FontSize(9).Bold();
                            txt.Span("Kapag hindi nabayaran o na-renew ang renta kahit lumampas na sa ").FontSize(9);
                            txt.Span("tatlong (3) buwang grace period mula sa expiration ng lease ").FontSize(9).Bold();
                            txt.Span(", ang cemetery lot ay maaaring ").FontSize(9);
                            txt.Span("ideklarang bakante").FontSize(9).Bold();
                            txt.Span("alinsunod sa umiiral na ordinansa.").FontSize(9);
                        });

                        col.Item().PaddingTop(2).Text(txt =>
                        {
                            txt.Span("Kung nabayaran na ang nasabing account, mangyaring magsumite ng kopya ng " +
                                     "Official Receipt upang ma-update at ma-reconcile ang aming talaan.")
                               .FontSize(9);
                        });

                        // Signature block
                        col.Item().PaddingTop(24).Text("Very truly yours,").FontSize(9);
                        col.Item().PaddingTop(30).Text("ALMA D. BUCALEN, CPA, MPA").FontSize(9).Bold();
                        col.Item().Text("Acting City Treasurer").FontSize(9);
                    });

                    // ── FOOTER ──
                    page.Footer().PaddingBottom(0.5f, Unit.Centimetre).Column(f =>
                    {
                        f.Item().LineHorizontal(1).LineColor(Colors.Red.Medium);
                        f.Item().PaddingTop(6).AlignCenter()
                            .Text("The Navotas City Hall M. Naval St., Sipac Almacen, Navotas City Zip: 1485")
                            .FontSize(9).Bold();
                        f.Item().AlignCenter().Text(txt =>
                        {
                            txt.Span("Tel. Nos. ").FontSize(9).Bold();
                            txt.Span("283-4210 Local 201, 202 & 813").FontSize(9).Bold();
                        });
                    });
                });
            }
        }).GeneratePdf(path);
    }
    private static bool IsZeroTotal(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return true;

        // Strip currency symbols, commas, spaces
        var cleaned = new string(value
            .Where(c => char.IsDigit(c) || c == '.' || c == '-')
            .ToArray());

        if (cleaned.Length == 0) return true;

        return decimal.TryParse(cleaned,
            System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture,
            out var n) && n == 0m;
    }

    private static void InfoLine(ColumnDescriptor col, string label, string value)
    {
        // Recognize "<number> <unit>" and separate them
        string mainValue = value ?? "-";
        string unit = "";

        var match = System.Text.RegularExpressions.Regex.Match(
            mainValue, @"^(\S+)\s+(.+)$");

        if (match.Success && mainValue.EndsWith("sq. m."))
        {
            mainValue = match.Groups[1].Value;
            unit = match.Groups[2].Value;
        }

        col.Item().Text(txt =>
        {
            txt.Span(label).FontSize(10).SemiBold();
            txt.Span(mainValue).FontSize(10).Underline();

            if (!string.IsNullOrEmpty(unit))
                txt.Span(" " + unit).FontSize(10);
        });
    }

    private static void ChargeRow(TableDescriptor table, string label, string value)
    {
        table.Cell()
            .Border(0.5f)
             .Padding(4)
             .AlignLeft().AlignMiddle()                 // ← vertically centered
             .Text(label).FontSize(9);

        table.Cell()
            .Border(0.5f)
             .Padding(4)
             .AlignRight().AlignMiddle()                // ← vertically centered
             .Text("₱ " + (string.IsNullOrWhiteSpace(value) ? "0.00" : value))
             .FontSize(9);
    }

    private static string MakeSafeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return string.IsNullOrWhiteSpace(name) ? Guid.NewGuid().ToString("N")[..8] : name;
    }

    private static bool AskYesNo(string msg) =>
        MessageBox.Show(msg, "Done", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes;

    private static void ShowError(string title, Exception ex) =>
        MessageBox.Show($"{title}:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

    private static void OpenFile(string path) =>
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });

    private void importTemplateToolStripMenuItem_Click(object sender, EventArgs e)
    {

    }

    private void OpenSettings()
    {
        using var dlg = new SettingsForm(
            Properties.Settings.Default.SkipZeroTotal,
            Properties.Settings.Default.PaperSize,
            Properties.Settings.Default.ExportFormat);

        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        // Save
        Properties.Settings.Default.SkipZeroTotal = dlg.SkipZeroTotal;
        Properties.Settings.Default.PaperSize = dlg.PaperSize;
        Properties.Settings.Default.ExportFormat = dlg.ExportFormat;

        Properties.Settings.Default.Save();

        lblStatus.Text = "Settings saved.";
        btnPrintPdf.Text = $"Export as {dlg.ExportFormat}";
        bool isPdf = dlg.ExportFormat == "PDF";
        fileToolStripMenuItem.Visible = !isPdf;
    }
    private static string GetAccountsCachePath()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "StatementOfAccountPrinting");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "accounts.json");
    }
    private void exportTemplateToolStripMenuItem_Click_1(object sender, EventArgs e)
    {
        ExportTemplate();
    }
    private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
    {
        OpenSettings();
    }

    #region BLOCK SELECTION...
    private void ShowBlockFilterPopup()
    {
        if (_accounts.Count == 0)
        {
            MessageBox.Show("Import an Excel file first.",
                "Block Filter", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (_blockPopup == null || _blockPopup.IsDisposed)
        {
            _blockPopup = new BlockFilterPopup();
            _blockPopup.SelectionChanged += OnBlockSelectionChanged;
            _blockPopup.VisibleChanged += OnBlockFilterVisibilityChanged;
        }

        // Distinct block names, sorted naturally
        var blocks = _accounts
            .Select(a => string.IsNullOrWhiteSpace(a.LotBlockNumber)
                            ? "(No Block)"
                            : a.LotBlockNumber.Trim())
            .Distinct()
            .OrderBy(b => ExtractBlockNumber(b))
            .ThenBy(b => b, StringComparer.OrdinalIgnoreCase)
            .ToList();

        IEnumerable<string>? preChecked = _blockFilterInitialized ? _selectedBlocks : null;

        _blockPopup.LoadBlocks(blocks, preChecked);
        _blockPopup.ShowUnder(btnBlockFilter);
    }
    private void OnBlockSelectionChanged(HashSet<string> selected)
    {
        _selectedBlocks.Clear();
        foreach (var b in selected)
            _selectedBlocks.Add(b);

        _blockFilterInitialized = true;   // user has interacted with the popup
        UpdateBlockFilterCaption();
    }
    private void OnBlockFilterVisibilityChanged(object? sender, EventArgs e)
    {
        // Only rebuild when the popup is becoming invisible
        if (_blockPopup != null && !_blockPopup.Visible)
            RebuildBlockTabs();
    }
    private void UpdateBlockFilterCaption()
    {
        int total = _accounts
            .Select(a => a.LotBlockNumber?.Trim() ?? "(No Block)")
            .Distinct()
            .Count();

        if (_selectedBlocks.Count == 0 || _selectedBlocks.Count == total)
            btnBlockFilter.Text = $"All Blocks ({total})";
        else if (_selectedBlocks.Count == 1)
            btnBlockFilter.Text = $"{_selectedBlocks.First()}";
        else
            btnBlockFilter.Text = $"{_selectedBlocks.Count} of {total} blocks";

        // Ensure it renders properly
        btnBlockFilter.AutoSize = false;
        btnBlockFilter.TextAlign = ContentAlignment.MiddleCenter;
        btnBlockFilter.Padding = new Padding(6, 2, 6, 2);
        // Optional: bump height if it looks cramped
        if (btnBlockFilter.Height < 28) btnBlockFilter.Height = 28;
    }
    private void btnBlockFilter_Click(object sender, EventArgs e)
    {
        ShowBlockFilterPopup();
    }
    #endregion
    #region SAVE LAST IMPORTED ACCOUNTS...
    private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
    {
        try
        {
            SaveAccountsToDisk();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine("Save failed: " + ex.Message);
        }
    }
    private void SaveAccountsToDisk()
    {
        var path = GetAccountsCachePath();

        // Include DateGenerated so the cached snapshot stays accurate
        var options = new System.Text.Json.JsonSerializerOptions
        {
            WriteIndented = true
        };

        var json = System.Text.Json.JsonSerializer.Serialize(_accounts.ToList(), options);
        File.WriteAllText(path, json);
    }
    private async Task LoadAccountsFromDiskAsync()
    {
        var path = GetAccountsCachePath();
        if (!File.Exists(path)) return;

        List<CemeteryLotAccount>? loaded = null;

        await ProgressDialog.RunAsync(
            this,
            "Loading saved records…",
            async progress =>
            {
                await Task.Run(() =>
                {
                    progress.Report((0, "Reading cache file…"));

                    try
                    {
                        var json = File.ReadAllText(path);
                        loaded = System.Text.Json.JsonSerializer
                            .Deserialize<List<CemeteryLotAccount>>(json);

                        progress.Report((50, $"Parsed {loaded?.Count ?? 0} record(s)."));
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("Load failed: " + ex.Message);
                        progress.Report((100, "Cache load failed — starting with an empty list."));
                        return;
                    }

                    progress.Report((100, "Finishing up…"));
                });
            },
            keepOpenWhenDone: false);   // auto-close — user doesn't need to dismiss it

        // ── Back on the UI thread ──
        if (loaded == null || loaded.Count == 0) return;

        _accounts.Clear();
        foreach (var a in loaded)
            _accounts.Add(a);

        RebuildBlockTabs();
        _selectedBlocks.Clear();
        UpdateBlockFilterCaption();

        btnGenerateWord.Enabled = btnPrintPdf.Enabled = _accounts.Count > 0;
        btnPrintSelected.Enabled = _accounts.Count > 0;
        lblStatus.Text = $"Loaded {_accounts.Count} cached record(s) from previous session.";

        if (btnPrintPdf.Enabled && btnPrintPdf.Visible)
        {
            btnBlockFilter.Enabled = true;
            btnPrintPdf.Focus();
            btnPrintPdf.Select();
        }
    }
    #endregion
    private void exitToolStripMenuItem_Click(object sender, EventArgs e)
    {
        CloseForm();
    }
    private void CloseForm()
    {
        var result = MessageBox.Show(
            "Close the application?",
            "Confirm",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result == DialogResult.Yes)
            Close();
    }
}