namespace StatementOfAccountPrinting
{
    public partial class SettingsForm : Form
    {
        public event Action? CacheCleared;
        public bool SkipZeroTotal { get; private set; } = false;
        public string PaperSize { get; private set; } = "A4";
        public string ExportFormat { get; private set; } = "PDF";

        public SettingsForm(bool skipZeroTotal, string paperSize, string exportFormat)
        {
            InitializeComponent();

            if (chkSkipZeroTotal != null)
                chkSkipZeroTotal.Checked = Properties.Settings.Default.SkipZeroTotal;

            var savedFormat = Properties.Settings.Default.ExportFormat;
            var formatIndex = cmbExportFormat.Items.IndexOf(savedFormat);
            cmbExportFormat.SelectedIndex = formatIndex >= 0 ? formatIndex : 0;   // default PDF

            if (cmbPaperSize != null)
            {
                var saved = Properties.Settings.Default.PaperSize;
                var idx = cmbPaperSize.Items.IndexOf(saved);
                cmbPaperSize.SelectedIndex = idx >= 0 ? idx : 0;
                //cmbPaperSize.Enabled = savedFormat == "PDF";
            }

            cmbPaperSize.SelectedItem =
                cmbPaperSize.Items.Contains(paperSize) ? paperSize : "A4";

            cmbExportFormat.SelectedItem =
                cmbExportFormat.Items.Contains(exportFormat) ? exportFormat : "PDF";

            //UpdateCacheStatus();
        }

        #region CLEAR CACHE...
        private void btnClearCache_Click(object sender, EventArgs e)
        {
            var confirm = MessageBox.Show(
                "Clear the cached imported records?\n\n" +
                "This does NOT modify any Excel file. The grid will be emptied " +
                "and the cache file will be deleted.",
                "Clear Cache",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes) return;

            try
            {
                var path = GetAccountsCachePath();
                if (File.Exists(path))
                    File.Delete(path);

                // Notify MainForm to clear its in-memory list and grid
                CacheCleared?.Invoke();

                MessageBox.Show("Cache cleared.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Could not clear the cache:\n\n" + ex.Message,
                    "Clear Cache",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        //private void UpdateCacheStatus()
        //{
        //    var path = GetAccountsCachePath();
        //    var message = File.Exists(path)
        //        ? $"Cache file: {path}"
        //        : "No cache file present.";

        //    MessageBox.Show(message);
        //}

        private static string GetAccountsCachePath()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "StatementOfAccountPrinting");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "accounts.json");
        }
        #endregion

        private void cmbExportFormat_SelectedIndexChanged(object sender, EventArgs e)
        {
            //cmbPaperSize.Enabled = cmbExportFormat.SelectedItem?.ToString() == "PDF";
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            SkipZeroTotal = chkSkipZeroTotal.Checked;
            PaperSize = cmbPaperSize.SelectedItem.ToString() ?? "A4";
            ExportFormat = cmbExportFormat.SelectedItem.ToString() ?? "PDF";
        }
    }
}
