using AutoUpdaterDotNET;
using ClosedXML.Excel;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Reflection;

namespace StatementOfAccountPrinting;

public partial class MainForm : Form
{
    private readonly BindingList<CemeteryLotAccount> _accounts = new();
    private readonly HashSet<string> _selectedBlocks = new(StringComparer.OrdinalIgnoreCase);
    private bool _blockFilterInitialized;
    private bool _isUpdating = false;
    private BlockFilterPopup? _blockPopup;
    private System.Windows.Forms.Timer _searchDebounceTimer;
    private readonly bool IsDevelopment =
#if DEBUG
        true;
#else
        false;
#endif

    public MainForm()
    {
        InitializeComponent();

        // Setup the debounce timer
        _searchDebounceTimer = new System.Windows.Forms.Timer();
        _searchDebounceTimer.Interval = 300; // 300 milliseconds
        _searchDebounceTimer.Tick += SearchDebounceTimer_Tick;

        LoadUIFromSettings();
        QuestPDF.Settings.License = LicenseType.Community;
    }
    private async void MainForm_Shown(object sender, EventArgs e)
    {
        await LoadAccountsFromDiskAsync();
    }
    private void CacheClearedUI()
    {
        tabBlocks.Visible = false;
        toolStripButtonImportExcel.Visible = false;

        var emptyStatePanel = new System.Windows.Forms.Panel
        {
            Name = "panelEmptyState",
            Width = 300,
            Height = 120,
            BackColor = System.Drawing.Color.Transparent
        };

        var messageLabel = new System.Windows.Forms.Label
        {
            Text = "No records found.\nImport an Excel file to get started.",
            Font = new Font("Segoe UI", 11F, FontStyle.Regular),
            ForeColor = System.Drawing.Color.Gray,
            TextAlign = ContentAlignment.MiddleCenter,
            Dock = DockStyle.Top,
            Height = 50
        };

        var importBtn = new System.Windows.Forms.Button
        {
            Text = "Import Excel",
            Width = 160,
            Height = 45,
            Font = new Font("Segoe UI", 11F, FontStyle.Regular),
            BackColor = System.Drawing.Color.White,
            FlatStyle = FlatStyle.System,
            Name = "btnEmptyStateImport",
            Image = Properties.Resources.sheets,
            TextImageRelation = TextImageRelation.ImageBeforeText
        };

        importBtn.Click += (s, e) => toolStripButtonImportExcel_Click(sender: null, e: EventArgs.Empty);

        emptyStatePanel.Controls.Add(importBtn);
        emptyStatePanel.Controls.Add(messageLabel);
        importBtn.Location = new Point(
            (emptyStatePanel.Width - importBtn.Width) / 2,
            messageLabel.Bottom + 10
        );
        void CenterPanel()
        {
            if (emptyStatePanel.IsDisposed) return;

            int centerX = tabBlocks.Left + (tabBlocks.Width - emptyStatePanel.Width) / 2;
            int centerY = tabBlocks.Top + (tabBlocks.Height - emptyStatePanel.Height) / 2;
            emptyStatePanel.Location = new Point(centerX, centerY);
        }

        CenterPanel();

        this.Resize += (s, e) => CenterPanel();
        tabBlocks.Resize += (s, e) => CenterPanel();

        this.Controls.Add(emptyStatePanel);
        emptyStatePanel.BringToFront();
    }
    private void RestoreTabsUI()
    {
        // 1) Remove the empty-state button if it exists
        var existingBtn = this.Controls.Find("btnEmptyStateImport", true).FirstOrDefault();
        if (existingBtn != null)
        {
            this.Controls.Remove(existingBtn);
            existingBtn.Dispose();
        }
        var emptyStatePanel = this.Controls.Find("panelEmptyState", true).FirstOrDefault();
        if (emptyStatePanel != null)
        {
            this.Controls.Remove(emptyStatePanel);
            emptyStatePanel.Dispose();
        }

        // 2) Show the tabs and toolbar button again
        tabBlocks.Visible = true;
        toolStripButtonImportExcel.Visible = true;
    }
    private void LoadUIFromSettings()
    {
        // 1. Try to get NBGV version (InformationalVersion)
        string displayVersion = Assembly.GetEntryAssembly()
            ?.GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion;

        // 2. If NBGV is empty, fallback to FileVersion (Windows Properties -> Details)
        if (string.IsNullOrEmpty(displayVersion))
        {
            try
            {
                string exePath = Assembly.GetEntryAssembly()?.Location;
                if (!string.IsNullOrEmpty(exePath) && File.Exists(exePath))
                    displayVersion = FileVersionInfo.GetVersionInfo(exePath).FileVersion;
            }
            catch { }
        }

        // 3. Final fallback if all else fails
        if (string.IsNullOrEmpty(displayVersion))
            displayVersion = "1.0.0.0";

        // (Optional) Remove the git hash for a cleaner title (e.g., "1.0.0.2+abc123" -> "1.0.0.2")
        string cleanVersion = displayVersion.Split('+')[0];

        string AppTitle = $"SoA Printing — Public Cemetery Lot Rental v{cleanVersion}";
        this.Text = AppTitle;

        RegisterFonts();

        var isWordExportEnabled = Properties.Settings.Default.EnableWordExport;
        var savedFormat = !isWordExportEnabled ? "PDF" : Properties.Settings.Default.ExportFormat;
        toolStripSplitButtonExport.Text = $"Export All as {savedFormat}";
        bool isPdf = savedFormat == "PDF";
        Bitmap bmp = isPdf ? Properties.Resources.pdf : Properties.Resources.doc;
        toolStripSplitButtonExport.Image = bmp;
        //fileToolStripMenuItem.Enabled = !isPdf;

        importTemplateToolStripMenuItem.Visible = isWordExportEnabled;
        exportTemplateToolStripMenuItem.Visible = isWordExportEnabled;
        toolStripSeparator1.Visible = isWordExportEnabled;

        string devSetting = ConfigurationManager.AppSettings["IsDevelopment"];
        //var IsDevelopment = bool.TryParse(devSetting, out bool result) && result;

        //checkForUpdatesToolStripMenuItem.Visible = IsDevelopment;
    }
    private async Task CheckForUpdatesBlocking()
    {
        AutoUpdater.SetOwner(this);
        AutoUpdater.RunUpdateAsAdmin = false;

        UpdateInfoEventArgs updateInfo = null;

        var tcs = new TaskCompletionSource<bool>();
        AutoUpdaterDotNET.AutoUpdater.CheckForUpdateEventHandler handler = null;

        handler = (args) =>
        {
            updateInfo = args;
            AutoUpdaterDotNET.AutoUpdater.CheckForUpdateEvent -= handler;
            tcs.TrySetResult(true);
        };

        AutoUpdaterDotNET.AutoUpdater.CheckForUpdateEvent += handler;
        AutoUpdaterDotNET.AutoUpdater.Start("https://raw.githubusercontent.com/ArJayEm/SOAPublicCemeteryLotRental/refs/heads/master/Updates/updates.xml");

        var completed = await Task.WhenAny(tcs.Task, Task.Delay(10000));
        if (completed != tcs.Task || updateInfo == null || updateInfo.Error != null)
        {
            MessageBox.Show("Could not connect to the update server.", "Update Error");
            return;
        }

        if (!updateInfo.IsUpdateAvailable)
        {
            MessageBox.Show("You are already on the latest version.", "No Updates");
            return;
        }

        // ── Ask user ──
        var result = MessageBox.Show(
            $"A new version ({updateInfo.CurrentVersion}) is available!\n\n" +
            "Do you want to download and install it now?\n\n" +
            "The app will briefly restart to apply the update.",
            "Update Available",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Information);

        if (result != DialogResult.Yes) return;

        // ── Show progress bar while downloading ──
        string tempInstaller = Path.Combine(Path.GetTempPath(), "soa-printing-cemetery.exe");

        bool downloadSuccess = false;

        await ProgressDialog.RunAsync(this, "Downloading update...", async prog =>
        {
            prog.Report((10, "Connecting..."));

            try
            {
                using var client = new System.Net.Http.HttpClient();
                using var response = await client.GetAsync(updateInfo.DownloadURL, HttpCompletionOption.ResponseHeadersRead);
                response.EnsureSuccessStatusCode();

                long? totalBytes = response.Content.Headers.ContentLength;
                using var contentStream = await response.Content.ReadAsStreamAsync();
                using var fileStream = new FileStream(tempInstaller, FileMode.Create, FileAccess.Write, FileShare.None);

                var buffer = new byte[8192];
                long totalRead = 0;
                int read;

                while ((read = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                {
                    await fileStream.WriteAsync(buffer, 0, read);
                    totalRead += read;

                    if (totalBytes.HasValue)
                    {
                        int pct = (int)((totalRead / (double)totalBytes.Value) * 100);
                        prog.Report((pct, $"Downloading... {totalRead / 1024} KB / {totalBytes.Value / 1024} KB"));
                    }
                }

                prog.Report((100, "Download complete."));
                downloadSuccess = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to download update:\n" + ex.Message, "Error");
            }
        }, keepOpenWhenDone: false);

        if (!downloadSuccess) return;

        // ── Confirm restart ──
        var restart = MessageBox.Show(
            "The update is ready.\n\n" +
            "The app will close now to install the update, then reopen automatically.\n\n" +
            "Continue?",
            "Restart Required",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (restart != DialogResult.Yes) return;

        // ── Silently run installer, close app, let Inno Setup reopen it ──
        try
        {
            // Inno Setup flags:
            //   /SILENT             = Show progress bar only, no wizard
            //   /CLOSEAPPLICATIONS  = Auto-close apps holding the target files
            //   /RESTARTAPPLICATIONS= Reopen the app after install
            //   /NORESTART          = Do not reboot Windows if files are locked
            string args = "/SILENT /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS /NORESTART";

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = tempInstaller,
                Arguments = args,
                UseShellExecute = true,
                Verb = "runas"   // UAC prompt for Program Files
            });

            // Force immediate exit so the installer can replace the .exe
            _isUpdating = true;
            Application.Exit();
        }
        catch (Exception ex)
        {
            MessageBox.Show("Failed to run installer:\n" + ex.Message, "Error");
        }
    }
    private void RegisterFonts()
    {
        List<Stream> _fontStreams = new();

        foreach (var path in new[]
        {
            "Resources/Fonts/calibri.ttf",
            "Resources/Fonts/calibrib.ttf",
            "Resources/Fonts/calibrii.ttf",
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
    private async Task RebuildBlockTabsAsync()
    {
        if (_accounts.Count == 0) return;

        string searchTerm = toolStripTextBoxSearch.Text.Trim();
        bool isSearching = !string.IsNullOrWhiteSpace(searchTerm);

        // 1. Filter the accounts
        var filteredAccounts = _accounts.AsEnumerable();

        if (!string.IsNullOrWhiteSpace(searchTerm))
        {
            filteredAccounts = filteredAccounts.Where(a =>
                (a.LesseeName != null && a.LesseeName.Contains(searchTerm, StringComparison.OrdinalIgnoreCase)) ||
                (a.NameOfDeceased != null && a.NameOfDeceased.Contains(searchTerm, StringComparison.OrdinalIgnoreCase)) ||
                (a.Address != null && a.Address.Contains(searchTerm, StringComparison.OrdinalIgnoreCase)) ||
                (a.LotBlockNumber != null && a.LotBlockNumber.Contains(searchTerm, StringComparison.OrdinalIgnoreCase))
            // Add any other fields you want to search here
            );
        }

        bool skipZero = Properties.Settings.Default.SkipZeroTotal;

        // Group first — cheap, can happen on UI thread
        var groups = filteredAccounts
            .Where(a => !skipZero || !IsZeroTotal(a.TotalAmountDue))
            .GroupBy(a =>
                // 👇 If searching, put everything in one "Search Results" group
                isSearching ? "Search Results" :
                (string.IsNullOrWhiteSpace(a.LotBlockNumber) ? "(No Block)" : a.LotBlockNumber.Trim())
            )
            .OrderBy(g => ExtractBlockNumber(g.Key))
            .ThenBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();

        // Apply the visible-block filter
        var allBlockNames = groups.Select(g => g.Key).ToList();
        bool showAll = _selectedBlocks.Count == 0
                       || _selectedBlocks.Count >= allBlockNames.Count;

        var visibleGroups = showAll
            ? groups
            : groups.Where(g => _selectedBlocks.Contains(g.Key)).ToList();

        // Preserve current tab
        string? current = tabBlocks.SelectedTab?.Tag as string;
        lblStatus.Text = "Rebuilding tabs…";

        bool ok = await ProgressDialog.RunAsync(this, "Rebuilding tabs…", async progress =>
        {
            await Task.Run(() =>
            {
                // Build each page on the background thread
                var builtPages = new List<TabPage>();

                for (int i = 0; i < visibleGroups.Count; i++)
                {
                    var group = visibleGroups[i];

                    progress.Report((
                        (int)((i / (double)visibleGroups.Count) * 100),
                        $"Building tab {i + 1} of {visibleGroups.Count} — {group.Key}"));

                    var page = new TabPage
                    {
                        Text = $"{group.Key} ({group.Count()})",
                        Tag = group.Key,
                        Padding = new Padding(4)
                    };

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
                        MultiSelect = true,
                        DataSource = blockList
                    };

                    // ✅ Friendly headers as columns are auto-generated
                    grid.ColumnAdded += (s, e) =>
                    {
                        var prop = e.Column.DataPropertyName;
                        if (string.IsNullOrEmpty(prop)) return;

                        e.Column.HeaderText = CemeteryLotAccount.ColumnHeaders
                            .TryGetValue(prop, out var label)
                                ? label
                                : SplitCamelCase(prop);
                    };

                    var menu = new ContextMenuStrip();
                    menu.Items.Add("Print Selected Record(s)", null,
                        (s, e) => toolStripRightClickExportSelected_Click(s, e));
                    grid.ContextMenuStrip = menu;

                    page.Controls.Add(grid);
                    builtPages.Add(page);
                }

                // Hand the finished pages back to the UI thread
                tabBlocks.Invoke(new Action(() =>
                {
                    tabBlocks.SuspendLayout();
                    try
                    {
                        tabBlocks.TabPages.Clear();
                        foreach (var p in builtPages)
                        {
                            // Re-hook selection handler now that the page is on the UI thread
                            var grid = p.Controls.OfType<DataGridView>().FirstOrDefault();
                            if (grid != null)
                            {
                                grid.SelectionChanged += (s, e) =>
                                    toolStripMenuItemExportSelected.Enabled = grid.SelectedRows.Count > 0;
                            }

                            tabBlocks.TabPages.Add(p);
                        }

                        // Restore previous tab if still present
                        if (current != null)
                        {
                            foreach (TabPage p in tabBlocks.TabPages)
                            {
                                if (string.Equals(p.Tag as string, current,
                                                  StringComparison.OrdinalIgnoreCase))
                                {
                                    tabBlocks.SelectedTab = p;
                                    break;
                                }
                            }
                        }

                        if (tabBlocks.TabPages.Count > 0 && tabBlocks.SelectedTab == null)
                            tabBlocks.SelectedTab = tabBlocks.TabPages[0];

                        toolStripMenuItemExportSelected.Enabled = tabBlocks.TabPages.Count > 0
                            && GetActiveGrid()?.SelectedRows.Count > 0;
                    }
                    finally
                    {
                        tabBlocks.ResumeLayout();
                    }
                }));

                progress.Report((100, "Done."));
            });
        }, keepOpenWhenDone: false);

        RestoreTabsUI();
    }
    private void TabBlocks_SelectedIndexChanged(object? sender, EventArgs e)
    {
        var g = GetActiveGrid();
        toolStripMenuItemExportSelected.Enabled = g?.SelectedRows.Count > 0;
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
    private async void toolStripMenuItemExportSelected_Click(object sender, EventArgs e)
    {
        PrintSelected();
    }
    private async void PrintSelected()
    {
        toolStripButtonImportExcel.Enabled = false;
        toolStripSplitButtonExport.Enabled = false;
        toolStripMenuItemExportSelected.Enabled = false;

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

            bool ok = await ProgressDialog.RunAsync(this,
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

            if (!ok)
            {
                lblStatus.Text = $"Export failed after {FormatElapsed(sw.Elapsed)}";
                return;   // ← no prompt
            }

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
            toolStripButtonImportExcel.Enabled = _accounts.Count > 0;
            toolStripSplitButtonExport.Enabled = _accounts.Count > 0;
            toolStripMenuItemExportSelected.Enabled = _accounts.Count > 0;
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
    private async void toolStripButtonImportExcel_Click(object sender, EventArgs e)
    {
        using var ofd = new OpenFileDialog
        {
            Title = "Select Excel file",
            Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        };
        if (ofd.ShowDialog() != DialogResult.OK) return;

        toolStripButtonImportExcel.Enabled = false;
        toolStripSplitButtonExport.Enabled = false;
        toolStripMenuItemBlockFilter.Enabled = false;
        //toolStripMenuItemExportSelected.Enabled = false;   // ← add this
        //toolStripMenuItemBlockFilter.Enabled = false;

        var filePath = ofd.FileName;
        var imported = new List<CemeteryLotAccount>();   // ← LOCAL list, not _accounts

        // ── Run the export with progress ──
        var sw = System.Diagnostics.Stopwatch.StartNew();

        bool ok = await ProgressDialog.RunAsync(this, "Importing Excel…", async progress =>
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

        sw.Stop();

        // ── Back on the UI thread — now safe to touch _accounts and dgvData ──
        _accounts.Clear();
        foreach (var a in imported)
            _accounts.Add(a);

        //PopulateBlockFilter();
        RebuildBlockTabsAsync();
        UpdateBlockFilterCaption();
        SaveAccountsToDisk();

        toolStripMenuItemExportSelected.Enabled = _accounts.Count > 0;   // ← add this
        lblStatus.Text = $"Imported {_accounts.Count} row(s) from " +
                         $"{Path.GetFileName(filePath)}.";

        if (toolStripSplitButtonExport.Enabled && toolStripSplitButtonExport.Visible)
        {
            toolStripMenuItemBlockFilter.Enabled = true;
            toolStripSplitButtonExport.Select();
        }

        toolStripButtonImportExcel.Enabled = _accounts.Count > 0;
        toolStripSplitButtonExport.Enabled = _accounts.Count > 0;
        toolStripMenuItemBlockFilter.Enabled = _accounts.Count > 0;
        //toolStripMenuItemExportSelected.Enabled = _accounts.Count > 0;
        //toolStripMenuItemBlockFilter.Enabled = _accounts.Count > 0;

        if (!ok)
        {
            lblStatus.Text = $"Export failed after {FormatElapsed(sw.Elapsed)}";
            return;   // ← no prompt
        }
    }
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
    private static string SplitCamelCase(string input)
    {
        if (string.IsNullOrEmpty(input)) return input ?? "";

        // Insert a space before each capital letter that follows a lowercase letter or digit
        var spaced = System.Text.RegularExpressions.Regex.Replace(
            input,
            @"(?<=[a-z0-9])([A-Z])",
            " $1");

        // Also handle digits glued to letters: "AreaSqm2" → "Area Sqm 2"
        spaced = System.Text.RegularExpressions.Regex.Replace(
            spaced,
            @"(?<=[A-Za-z])(\d)",
            " $1");

        return spaced.Trim();
    }
    #endregion
    private void btnCreateTemplate_Click(object sender, EventArgs e)
    {
        ExportDocxTemplate();
    }
    private void ExportDocxTemplate()
    {
        using var sfd = new SaveFileDialog
        {
            InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            Title = "Save blank template as",
            FileName = WordTemplateService.TemplateFileName,
            DefaultExt = "docx",
            Filter = "Word documents (*.docx)|*.docx|All files (*.*)|*.*",
            AddExtension = true,
            OverwritePrompt = true
        };

        if (sfd.ShowDialog() != DialogResult.OK) return;

        //exportTemplateToolStripMenuItem.Enabled = false;
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
            //exportTemplateToolStripMenuItem.Enabled = true;
        }
    }
    private async Task ExportAsWord()
    {
        using var sfd = new SaveFileDialog
        {
            InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
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
    private async void toolStripRightClickExportSelected_Click(object sender, EventArgs e)
    {
        PrintSelected();
    }
    private async Task ExportAllAsPdfAsync(bool skipZero = false, bool onlyZero = false)
    {
        if (_accounts.Count == 0)
        {
            MessageBox.Show("No accounts to export.", "PDF",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        string paperSize = Properties.Settings.Default.PaperSize ?? "A4";
        //bool skipZero = Properties.Settings.Default.SkipZeroTotal;

        // ── Build the filtered source as a List<CemeteryLotAccount> ──
        List<CemeteryLotAccount> source = _accounts.ToList();

        using var fbd = new FolderBrowserDialog
        {
            Description = "Choose output folder for per-block PDFs"
        };
        if (fbd.ShowDialog() != DialogResult.OK) return;

        //toolStripButtonImportExcel.Enabled = false;
        //toolStripSplitButtonExport.Enabled = false;
        //toolStripMenuItemBlockFilter.Enabled = false;
        //toolStripMenuItemExportSelected.Enabled = false;
        //toolStripMenuItemBlockFilter.Enabled = false;

        var outputFolder = fbd.SelectedPath;

        // ── Read selected blocks from the popup's stored set ──
        var selectedBlocks = new HashSet<string>(_selectedBlocks, StringComparer.OrdinalIgnoreCase);

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
        // 2.5) Only-zero filter
        else if (onlyZero)
            source = source.Where(a => IsZeroTotal(a.TotalAmountDue)).ToList();

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

        var sw = System.Diagnostics.Stopwatch.StartNew();

        bool ok = await ProgressDialog.RunAsync(this, "Generating PDFs…", async progress =>
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

        // Reset block filter after export
        //_selectedBlocks.Clear();
        //_blockFilterInitialized = false;
        //UpdateBlockFilterCaption();

        sw.Stop();

        //toolStripButtonImportExcel.Enabled = _accounts.Count > 0;
        //toolStripSplitButtonExport.Enabled = _accounts.Count > 0;
        //toolStripMenuItemBlockFilter.Enabled = _accounts.Count > 0;
        //toolStripMenuItemExportSelected.Enabled = _accounts.Count > 0;
        //toolStripMenuItemBlockFilter.Enabled = _accounts.Count > 0;

        if (!ok)
        {
            lblStatus.Text = $"PDF creation failed after {FormatElapsed(sw.Elapsed)}";
            return;   // ← no prompt
        }

        lblStatus.Text = $"Saved {totalBlocks} block PDF(s) in {FormatElapsed(sw.Elapsed)} → {outputFolder}";

        if (AskYesNo("Open the output folder?"))
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            { FileName = outputFolder, UseShellExecute = true });
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
        var logoLeft = Path.Combine(baseDir, "Resources", "Images", "navotas-logo-left.jpg");
        var logoRight = Path.Combine(baseDir, "Resources", "Images", "blgf-logo-right.png");

        byte[]? logoLeftBytes = File.Exists(logoLeft) ? File.ReadAllBytes(logoLeft) : null;
        byte[]? logoRightBytes = File.Exists(logoRight) ? File.ReadAllBytes(logoRight) : null;

        Document.Create(doc =>
        {
            foreach (var a in list)          // ← only the accounts passed in
            {
                var dateGenerated = DateTime.Today.ToString("MMMM dd, yyyy");

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
                                InfoLine(inf, "Date / Petsa: ", dateGenerated);
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
        })
        .WithSettings(new DocumentSettings
        {
            ImageCompressionQuality = ImageCompressionQuality.Best,
            ImageRasterDpi = 150
        })
        .GeneratePdf(path);
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
        var isWordExportEnabled = Properties.Settings.Default.EnableWordExport;
        var skipZeroTotal = Properties.Settings.Default.SkipZeroTotal;
        var savedFormat = !isWordExportEnabled ? "PDF" : Properties.Settings.Default.ExportFormat;

        using var dlg = new SettingsForm(
            this,
            skipZeroTotal,
            Properties.Settings.Default.PaperSize,
            savedFormat);

        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        // Save
        Properties.Settings.Default.SkipZeroTotal = dlg.SkipZeroTotal;
        Properties.Settings.Default.PaperSize = dlg.PaperSize;
        Properties.Settings.Default.ExportFormat = dlg.ExportFormat;

        Properties.Settings.Default.Save();

        lblStatus.Text = "Settings saved.";
        toolStripSplitButtonExport.Text = $"Export All as {dlg.ExportFormat}";
        bool isPdf = dlg.ExportFormat == "PDF";
        Bitmap bmp = isPdf ? Properties.Resources.pdf : Properties.Resources.doc;
        toolStripSplitButtonExport.Image = bmp;

        if (skipZeroTotal != dlg.SkipZeroTotal)
            RebuildBlockTabsAsync();

        //fileToolStripMenuItem.Enabled = !isPdf;
    }
    public static string GetAccountsCachePath()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "StatementOfAccountPrinting");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "accounts.json");
    }
    private void exportTemplateToolStripMenuItem_Click_1(object sender, EventArgs e)
    {
        ExportDocxTemplate();
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

        // Group once, then emit (name, count) pairs in natural order
        var blocksWithCounts = _accounts
            .GroupBy(a => string.IsNullOrWhiteSpace(a.LotBlockNumber)
                            ? "(No Block)"
                            : a.LotBlockNumber.Trim())
            .Select(g => (Name: g.Key, Count: g.Count()))
            .OrderBy(x => ExtractBlockNumber(x.Name))
            .ThenBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        IEnumerable<string>? preChecked = _blockFilterInitialized
            ? _selectedBlocks
            : null;

        _blockPopup.LoadBlocks(blocksWithCounts, preChecked);
        _blockPopup.ShowUnder(toolStripMenuItemBlockFilter);
    }
    private void OnBlockFilterVisibilityChanged(object? sender, EventArgs e)
    {
        if (_blockPopup != null && !_blockPopup.Visible)
            RebuildBlockTabsAsync();
    }
    private void OnBlockSelectionChanged(HashSet<string> selected)
    {
        _selectedBlocks.Clear();
        foreach (var b in selected)
            _selectedBlocks.Add(b);

        _blockFilterInitialized = true;
        UpdateBlockFilterCaption();
    }
    private void UpdateBlockFilterCaption()
    {
        int total = _accounts
            .Select(a => a.LotBlockNumber?.Trim() ?? "(No Block)")
            .Distinct()
            .Count();

        if (_selectedBlocks.Count == 0 || _selectedBlocks.Count == total)
            toolStripMenuItemBlockFilter.Text = $"All Blocks ({total})";
        else if (_selectedBlocks.Count == 1)
            toolStripMenuItemBlockFilter.Text = $"{_selectedBlocks.First()}";
        else
            toolStripMenuItemBlockFilter.Text = $"{_selectedBlocks.Count} of {total} blocks";

        // Ensure it renders properly
        toolStripMenuItemBlockFilter.AutoSize = false;
        toolStripMenuItemBlockFilter.TextAlign = ContentAlignment.MiddleCenter;
        toolStripMenuItemBlockFilter.Padding = new Padding(6, 2, 6, 2);
        // Optional: bump height if it looks cramped
        if (toolStripMenuItemBlockFilter.Height < 28) toolStripMenuItemBlockFilter.Height = 28;
    }
    private void toolStripMenuItemBlockFilter_Click(object sender, EventArgs e)
    {
        ShowBlockFilterPopup();
    }
    #endregion
    #region SAVE LAST IMPORTED ACCOUNTS...
    private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
    {
        // 1. If we are restarting for an update, save and let it close silently
        if (_isUpdating)
        {
            try { SaveAccountsToDisk(); } catch { }
            return; // e.Cancel is false, so it closes.
        }

        // 2. If the user is trying to close the app normally (X button, menu, etc.)
        var result = MessageBox.Show(
            "Close the application?",
            "Confirm",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        // If they click NO, we MUST set e.Cancel = true to stop the close!
        if (result != DialogResult.Yes)
        {
            e.Cancel = true;
            return;
        }

        // 3. If they click YES, save data. The form will close automatically.
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

        var options = new System.Text.Json.JsonSerializerOptions
        {
            WriteIndented = true
        };

        var json = System.Text.Json.JsonSerializer.Serialize(_accounts.ToList(), options);
        File.WriteAllText(path, json);
    }
    public void ClearCachedRecords()
    {
        _accounts.Clear();
        _selectedBlocks.Clear();
        _blockFilterInitialized = false;

        RebuildBlockTabsAsync();
        UpdateBlockFilterCaption();
        CacheClearedUI();

        toolStripMenuItemBlockFilter.Enabled = false;
        toolStripSplitButtonExport.Enabled = false;

        lblStatus.Text = "Cached records cleared.";
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
        if (loaded == null || loaded.Count == 0)
        {
            CacheClearedUI();
            return;
        }

        _accounts.Clear();
        foreach (var a in loaded)
            _accounts.Add(a);

        RebuildBlockTabsAsync();
        _selectedBlocks.Clear();
        UpdateBlockFilterCaption();

        toolStripButtonImportExcel.Enabled = _accounts.Count > 0;
        toolStripMenuItemBlockFilter.Enabled = _accounts.Count > 0;
        toolStripSplitButtonExport.Enabled = _accounts.Count > 0;

        bool skipZero = Properties.Settings.Default.SkipZeroTotal;

        lblStatus.Text = $"Loaded {_accounts.Count} cached record(s) from previous session.{(skipZero ? " (skipped ₱0.00)" : "")}";

        if (toolStripSplitButtonExport.Enabled && toolStripSplitButtonExport.Visible)
        {
            toolStripMenuItemBlockFilter.Enabled = true;
            toolStripSplitButtonExport.Select();
        }
    }
    #endregion
    private void exitToolStripMenuItem_Click(object sender, EventArgs e)
    {
        this.Close();
    }
    private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
    {
        using var dlg = new About();
        if (dlg.ShowDialog(this) != DialogResult.OK) return;
    }

    private void toolStripSplitButtonExport_Click(object sender, EventArgs e)
    {
    }

    private void toolStripSplitButtonExport_ButtonClick(object sender, EventArgs e)
    {
        ExportAllAsPdfAsync();
    }

    private void exportAllskip000TotalAmountDueToolStripMenuItem_Click(object sender, EventArgs e)
    {
        ExportAllAsPdfAsync(skipZero: true);
    }

    private void exportAllonly000TotalAmountDueToolStripMenuItem_Click(object sender, EventArgs e)
    {
        ExportAllAsPdfAsync(onlyZero: true);
    }

    private async void checkForUpdatesToolStripMenuItem_Click(object sender, EventArgs e)
    {
        await CheckForUpdatesBlocking();
    }

    private void toolStripTextBoxSearch_KeyDown(object sender, KeyEventArgs e)
    {
        // Check if the user pressed Enter
        //if (e.KeyCode == Keys.Enter)
        //{
        //    e.SuppressKeyPress = true;
        //    e.Handled = true;

        //    _ = RebuildBlockTabsAsync();
        //}
    }
    private void toolStripTextBoxSearch_TextChanged(object sender, EventArgs e)
    {
        _searchDebounceTimer.Stop();
        _searchDebounceTimer.Start();

        string searchTerm = toolStripTextBoxSearch.Text.Trim();
        toolStripButtonClear.Enabled = !string.IsNullOrEmpty(searchTerm);
    }
    private void SearchDebounceTimer_Tick(object sender, EventArgs e)
    {
        _searchDebounceTimer.Stop();
        _ = RebuildBlockTabsAsync();
    }

    private void toolStripButtonClear_Click(object sender, EventArgs e)
    {
        toolStripTextBoxSearch.Text = "";
    }
}