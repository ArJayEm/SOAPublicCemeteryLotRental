using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.ComponentModel;
using System.Drawing.Printing;
using System.Text.RegularExpressions;

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

    public MainForm()
    {
        InitializeComponent();

        FormClosing += MainForm_FormClosing;

        Text = AppTitle;
        //dgvData.DataSource = _accounts;
        //RebuildBlockTabs();
        QuestPDF.Settings.License = LicenseType.Community;
        RegisterFonts();

        ToolTip toolTip = new ToolTip();
        toolTip.SetToolTip(btnBlockFilter, "Select Blocks to Print");
        toolTip.SetToolTip(btnPrintPdf, "Export as PDF");
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

        // Group by block
        var groups = _accounts
            .GroupBy(a => string.IsNullOrWhiteSpace(a.LotBlockNumber)
                            ? "(No Block)"
                            : a.LotBlockNumber.Trim())
            //.OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
            .OrderBy(g => ExtractBlockNumber(g.Key))         // 1, 2, 10, 100, …
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
                DataSource = blockList
            };

            page.Controls.Add(grid);
            tabBlocks.TabPages.Add(page);
        }

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

        // If the "All" tab was requested, keep it as the first tab (see below)
        //PrependAllTab();
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
            DataSource = _accounts   // the master BindingList
        };

        allPage.Controls.Add(grid);
        tabBlocks.TabPages.Insert(0, allPage);
    }

    #region IMPORT...
    private async void btnImportExcel_Click(object sender, EventArgs e)
    {
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
        }
    }
    //private void PopulateBlockFilter()
    //{
    //    clbBlockFilter.Items.Clear();

    //    var blocks = _accounts
    //        .Select(a => string.IsNullOrWhiteSpace(a.LotBlockNumber)
    //                        ? "(No Block)"
    //                        : a.LotBlockNumber.Trim())
    //        .Distinct()
    //        .OrderBy(b => ExtractBlockNumber(b))
    //        .ThenBy(b => b, StringComparer.OrdinalIgnoreCase)
    //        .ToList();

    //    foreach (var block in blocks)
    //    {
    //        // Checked by default = export all
    //        clbBlockFilter.Items.Add(block, true);
    //    }
    //}
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

    // ─────────────────────────────────────────────────────────────
    // Fetch cell text safely by column number.
    // Handles missing cells, dates, numbers, and trims whitespace.
    // ─────────────────────────────────────────────────────────────
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
            WordTemplateService.GenerateCombinedWord(sfd.FileName, _accounts.ToList());

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

    // ────────────── 4. Print all rows to one PDF ──────────────
    private async void btnPrintPdf_Click(object sender, EventArgs e)
    {
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
            UpdateBlockFilterCaption();
        }
        catch (Exception ex)
        {
            sw.Stop();
            ShowError($"PDF creation failed after {FormatElapsed(sw.Elapsed)}", ex);
        }
        finally
        {
            btnImportExcel.Enabled = btnGenerateWord.Enabled = btnPrintPdf.Enabled = true;
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
    /// <summary>
    /// True if the string represents a zero, empty, or unparseable amount.
    /// Handles "0", "0.00", "₱ 0.00", "-", "", null.
    /// </summary>
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
    private void CreatePdfByBlock(string path, string paperSize)
    {
        if (_accounts.Count == 0)
        {
            MessageBox.Show("No accounts to export.", "PDF",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // Group accounts by block / lot number
        var groups = _accounts
            .GroupBy(a => string.IsNullOrWhiteSpace(a.LotBlockNumber)
                            ? "(No Block)"
                            : a.LotBlockNumber.Trim())
            .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();

        // Logos
        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        var logoLeft = Path.Combine(baseDir, "Templates", "Images", "navotas-logo-left.jpg");
        var logoRight = Path.Combine(baseDir, "Templates", "Images", "blgf-logo-right.png");

        byte[]? logoLeftBytes = File.Exists(logoLeft) ? File.ReadAllBytes(logoLeft) : null;
        byte[]? logoRightBytes = File.Exists(logoRight) ? File.ReadAllBytes(logoRight) : null;

        Document.Create(doc =>
        {
            foreach (var group in groups)
            {
                // ── One page per block ──
                doc.Page(page =>
                {
                    ApplyPageSetup(page, paperSize);
                    DrawHeader(page, logoLeftBytes, logoRightBytes);
                    DrawBlockContent(page, group.Key, group.ToList());
                    DrawFooter(page);
                });
            }
        }).GeneratePdf(path);

        MessageBox.Show(
            $"PDF generated for {groups.Count} block(s), {_accounts.Count} record(s) ({paperSize}).",
            "PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private static void ApplyPageSetup(PageDescriptor page, string paperSize)
    {
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
    }

    private static void DrawHeader(PageDescriptor page, byte[]? logoLeft, byte[]? logoRight)
    {
        page.Header().PaddingTop(0).Column(h =>
        {
            h.Item().Row(row =>
            {
                row.ConstantItem(65).Height(65).AlignLeft().AlignMiddle().Element(e =>
                {
                    if (logoLeft != null) e.Image(logoLeft).FitArea();
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
                    if (logoRight != null) e.Image(logoRight).FitArea();
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
    }

    private static void DrawFooter(PageDescriptor page)
    {
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
    }

    private static void DrawBlockContent(PageDescriptor page, string blockName, List<CemeteryLotAccount> rows)
    {
        page.Content().Column(col =>
        {
            col.Spacing(0);

            // ── Block Title ──
            col.Item().PaddingTop(4).AlignCenter()
                .Text($"BLOCK / LOT: {blockName}")
                .FontSize(11).Bold();

            col.Item().PaddingTop(10);

            // ── Per-row table ──
            col.Item().Table(table =>
            {
                table.ColumnsDefinition(c =>
                {
                    c.RelativeColumn(3);   // name / address / deceased
                    c.RelativeColumn(1);   // area
                    c.RelativeColumn(1);   // rental
                    c.RelativeColumn(1);   // surcharge
                    c.RelativeColumn(1);   // interest
                    c.RelativeColumn(1);   // total
                });

                // Header row
                HeaderCell(table, "NAME / ADDRESS");
                HeaderCell(table, "AREA");
                HeaderCell(table, "RENTAL");
                HeaderCell(table, "SURCHARGE");
                HeaderCell(table, "INTEREST");
                HeaderCell(table, "TOTAL");

                // Data rows
                decimal sumRental = 0m;
                decimal sumSurcharge = 0m;
                decimal sumInterest = 0m;
                decimal sumTotal = 0m;

                foreach (var r in rows)
                {
                    decimal.TryParse(r.RentalOrRenewalFee, out var rental);
                    decimal.TryParse(r.Surcharge, out var surcharge);
                    decimal.TryParse(r.Interest, out var interest);
                    decimal.TryParse(r.TotalAmountDue, out var total);

                    sumRental += rental;
                    sumSurcharge += surcharge;
                    sumInterest += interest;
                    sumTotal += total;

                    DataCell(table,
                        $"{r.LesseeName}\n{r.Address}\nDeceased: {r.NameOfDeceased}",
                        left: true);

                    DataCell(table, r.AreaSqm + " sq.m.");
                    DataCell(table, "₱ " + rental.ToString("N2"));
                    DataCell(table, "₱ " + surcharge.ToString("N2"));
                    DataCell(table, "₱ " + interest.ToString("N2"));
                    DataCell(table, "₱ " + total.ToString("N2"));
                }

                // Footer totals row
                TotalsCell(table, $"TOTAL — {rows.Count} record(s)");
                TotalsCell(table, "");
                TotalsCell(table, "₱ " + sumRental.ToString("N2"));
                TotalsCell(table, "₱ " + sumSurcharge.ToString("N2"));
                TotalsCell(table, "₱ " + sumInterest.ToString("N2"));
                TotalsCell(table, "₱ " + sumTotal.ToString("N2"));
            });

            // Signature block
            col.Item().PaddingTop(24).Text("Very truly yours,").FontSize(9);
            col.Item().PaddingTop(30).Text("ALMA D. BUCALEN, CPA, MPA").FontSize(9).Bold();
            col.Item().Text("Acting City Treasurer").FontSize(9);
        });
    }

    private static void HeaderCell(TableDescriptor table, string text)
    {
        table.Cell()
            .Border(0.5f).Padding(3)
            .AlignCenter().AlignMiddle()
            .Text(text).FontSize(9).Bold();
    }

    private static void DataCell(TableDescriptor table, string text, bool left = false)
    {
        var cell = table.Cell().Border(0.5f).Padding(3);
        var aligned = left ? cell.AlignLeft() : cell.AlignRight();
        aligned.AlignMiddle().Text(text).FontSize(9);
    }

    private static void TotalsCell(TableDescriptor table, string text)
    {
        table.Cell()
            .Border(0.5f).Padding(3)
            .Background(Colors.Grey.Lighten3)
            .AlignRight().AlignMiddle()
            .Text(text).FontSize(9).Bold();
    }

    private void CreatePdf(string path, string paperSize)
    {
        if (_accounts.Count == 0)
        {
            MessageBox.Show("No accounts to export.", "PDF",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // Resolve logo paths (adjust folder if needed)
        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        var logoLeft = Path.Combine(baseDir, "Templates", "Images", "navotas-logo-left.jpg");
        var logoRight = Path.Combine(baseDir, "Templates", "Images", "blgf-logo-right.png");

        byte[]? logoLeftBytes = File.Exists(logoLeft) ? File.ReadAllBytes(logoLeft) : null;
        byte[]? logoRightBytes = File.Exists(logoRight) ? File.ReadAllBytes(logoRight) : null;

        Document.Create(doc =>
        {
            foreach (var a in _accounts)
            {
                doc.Page(page =>
                {
                    // ── A4, standard 1-inch margins ──

                    switch (paperSize)
                    {
                        case "Legal":
                            page.Size(PageSizes.Legal);   // 8.5 × 14 in
                            break;
                        case "A4":
                        default:
                            page.Size(PageSizes.A4);      // 8.27 × 11.69 in
                            break;
                    }

                    page.MarginTop(1, Unit.Centimetre);
                    page.MarginLeft(2.54f, Unit.Centimetre);
                    page.MarginRight(2.54f, Unit.Centimetre);
                    page.MarginBottom(0.5f, Unit.Centimetre);
                    page.DefaultTextStyle(t => t
                        .FontFamily("Lato")
                        .FontSize(9)
                        .FontColor(Colors.Black));

                    // ═══════════════════════════════════════════════
                    // HEADER — two logos + office titles + rule
                    // ═══════════════════════════════════════════════
                    page.Header().PaddingTop(0).Column(h =>
                    {
                        // Logos row
                        h.Item().Row(row =>
                        {
                            h.Item().Row(row =>
                            {
                                row.ConstantItem(70)
                                   .Height(70)
                                   .AlignLeft().AlignMiddle()
                                   .Element(e =>
                                   {
                                       if (logoLeftBytes != null)
                                           e.Image(logoLeftBytes).FitArea();
                                   });

                                row.RelativeItem().AlignCenter().AlignMiddle().Column(center =>
                                {
                                    center.Item().AlignCenter()
                                        .Text("Republic of the Philippines")
                                        .FontSize(14)
                                        .FontFamily("Times New Roman");

                                    center.Item().AlignCenter()
                                        .Text("CITY OF NAVOTAS")
                                        .FontSize(14).Bold()
                                        .FontFamily("Times New Roman")
                                        .FontColor(Colors.Grey.Darken3);

                                    center.Item().AlignCenter()
                                        .Text("OFFICE OF THE CITY TREASURER")
                                        .FontSize(14)
                                        .FontFamily("Times New Roman");
                                });

                                row.ConstantItem(70)
                                   .Height(70)
                                   .AlignRight().AlignMiddle()
                                   .Element(e =>
                                   {
                                       if (logoRightBytes != null)
                                           e.Image(logoRightBytes).FitArea();
                                   });
                            });

                        });

                        // Title
                        h.Item().PaddingTop(20).AlignCenter()
                            .Text("STATEMENT OF ACCOUNT – PUBLIC CEMETERY LOT RENTAL")
                            .FontSize(9).Bold();

                        h.Item().PaddingTop(2).AlignCenter()
                            .Text("(Pahayag ng Bayarin sa Renta ng Lupa sa Navotas City Public Cemetery)")
                            .FontSize(9).Bold();

                        h.Item().PaddingTop(14);
                    });

                    // ═══════════════════════════════════════════════
                    // CONTENT
                    // ═══════════════════════════════════════════════
                    page.Content().Column(col =>
                    {
                        col.Spacing(0);

                        // ── Info block (with checkbox to the right) ──
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

                        // ── Intro paragraph ──
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

                        // ── Charges Table
                        col.Item()
                           .PaddingTop(12)
                           .AlignCenter()          // ← centers the table horizontally in the content column
                           .Width(300)             // ← table width
                           .Table(table =>
                           {
                               table.ColumnsDefinition(c =>
                               {
                                   c.RelativeColumn(3);
                                   c.RelativeColumn(2);
                               });

                               // Header row
                               table.Cell()
                                .Border(0.5f).Padding(4)
                                .AlignCenter().AlignMiddle()
                                 .Text("PARTICULARS / DETALYE").FontSize(9).Bold();

                               table.Cell()
                               .Border(0.5f).Padding(4)
                                .AlignCenter().AlignMiddle()
                                .Text("AMOUNT / HALAGA").FontSize(9).Bold();

                               // Data rows
                               ChargeRow(table, "Rental / Renewal Fee", a.RentalOrRenewalFee);
                               ChargeRow(table, "Surcharge, if applicable", a.Surcharge);
                               ChargeRow(table, "Interest, if applicable", a.Interest);

                               // Total row (label has red italic subtitle)
                               table.Cell()
               .Border(0.5f)
                               .AlignLeft().AlignMiddle().Padding(4).Element(cell =>
                               {
                                   cell.Text(txt =>
                                   {
                                       txt.Span("Total Amount Due / ").FontSize(9);
                                       txt.Span("Kabuuang Babayaran").FontSize(9).Italic();
                                   });
                               });
                               table.Cell()
               .Border(0.5f).Padding(4)
                                    .AlignRight()
                                    .Text("₱ " + a.TotalAmountDue).FontSize(9);
                           });

                        // ── Payment and Inquiry Instructions ──
                        col.Item().PaddingTop(24).AlignCenter()
                            .Text("PAYMENT AND INQUIRY INSTRUCTIONS")
                            .FontSize(9).Bold();

                        col.Item().PaddingTop(2).AlignCenter()
                            .Text("(Paraan ng Pagbabayad at Pagtatanong)")
                            .FontSize(9).Bold();

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
                            txt.Span("ideklarang bakante ").FontSize(9).Bold();
                            txt.Span("alinsunod sa umiiral na ordinansa.").FontSize(9);
                        });

                        col.Item().PaddingTop(2).Text(txt =>
                        {
                            txt.Span("Kung nabayaran na ang nasabing account, mangyaring magsumite ng kopya ng " +
                                     "Official Receipt upang ma-update at ma-reconcile ang aming talaan.")
                               .FontSize(9);
                        });

                        // ── Signature Block ──
                        col.Item().PaddingTop(24).Text("Very truly yours,").FontSize(9);
                        col.Item().PaddingTop(30).Text("ALMA D. BUCALEN, CPA, MPA").FontSize(9).Bold();
                        col.Item().Text("Acting City Treasurer").FontSize(9);
                    });

                    // ═══════════════════════════════════════════════
                    // FOOTER — red rule + two-line contact info
                    // ═══════════════════════════════════════════════
                    page.Footer().PaddingBottom(0, Unit.Centimetre).Column(f =>
                    {
                        // Red rule
                        f.Item().LineHorizontal(1).LineColor(Colors.Red.Medium);

                        // Address
                        f.Item().PaddingTop(6).AlignCenter()
                            .Text("The Navotas City Hall M. Naval St., Sipac Almacen, Navotas City Zip: 1485")
                            .FontSize(9).Bold();

                        // Phone
                        f.Item().AlignCenter().Text(txt =>
                        {
                            txt.Span("Tel. Nos. ").FontSize(9).Bold();
                            txt.Span("283-4210 Local 201, 202 & 813").FontSize(9).Bold();
                        });
                    });
                });
            }
        }).GeneratePdf(path);

        BeginInvoke(new Action(() =>
        {
            MessageBox.Show(
                $"PDF generated for {_accounts.Count} of {_accounts.Count} account(s).",
                "PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }));
    }

    // One A4 page per row — page count == row count
    public static class SoaTemplate
    {
        // Match these exactly with your Templates/soa-public_cemetery_lot_rental.docx
        public const string HeaderTitle = "STATEMENT OF ACCOUNT – PUBLIC CEMETERY LOT RENTAL";
        public const string HeaderSubtitle = "(Pahayag ng Bayarin sa Renta ng Lupa sa Navotas City Public Cemetery)";

        public const string FooterLeft = "Navotas City Public Cemetery";
        public const string FooterRight = "City Treasurer’s Office";
    }

    // ─────────────────────────────────────────────
    // Info line (label semibold + value regular)
    // ─────────────────────────────────────────────
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

    // ─────────────────────────────────────────────
    // Charges row helper
    // ─────────────────────────────────────────────
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

    private string? PickTemplate()
    {
        // Offer the app-folder template first, else browse for one
        string local = Path.Combine(AppContext.BaseDirectory, TemplateFileName);
        if (File.Exists(local)) return local;

        using var ofd = new OpenFileDialog
        { Title = "Select the Word template", Filter = "Word documents (*.docx)|*.docx" };
        return ofd.ShowDialog() == DialogResult.OK ? ofd.FileName : null;
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
            Properties.Settings.Default.PaperSize);

        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        // Save
        Properties.Settings.Default.SkipZeroTotal = dlg.SkipZeroTotal;
        Properties.Settings.Default.PaperSize = dlg.PaperSize;
        Properties.Settings.Default.Save();

        lblStatus.Text = "Settings saved.";
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

        // If nothing is stored yet, default to "all checked"
        IEnumerable<string>? preChecked = _selectedBlocks.Count == 0
            ? null
            : _selectedBlocks;

        _blockPopup.LoadBlocks(blocks, preChecked);
        _blockPopup.ShowUnder(btnBlockFilter);
    }
    private void OnBlockSelectionChanged(HashSet<string> selected)
    {
        _selectedBlocks.Clear();
        foreach (var b in selected)
            _selectedBlocks.Add(b);

        UpdateBlockFilterCaption();
    }
    private void UpdateBlockFilterCaption()
    {
        int total = _accounts
            .Select(a => a.LotBlockNumber?.Trim() ?? "(No Block)")
            .Distinct()
            .Count();

        if (_selectedBlocks.Count == 0 || _selectedBlocks.Count == total)
            btnBlockFilter.Text = $"All Blocks ({total}) ▾";
        else if (_selectedBlocks.Count == 1)
            btnBlockFilter.Text = $"{_selectedBlocks.First()} ▾";
        else
            btnBlockFilter.Text = $"{_selectedBlocks.Count} of {total} blocks ▾";
    }
    #endregion

    private void btnBlockFilter_Click(object sender, EventArgs e)
    {
        ShowBlockFilterPopup();
    }
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