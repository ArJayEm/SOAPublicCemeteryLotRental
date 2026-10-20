namespace StatementOfAccountPrinting
{
    public partial class SettingsForm : Form
    {
        private readonly MainForm? _mainForm;
        public bool SkipZeroTotal { get; private set; } = false;
        public string PaperSize { get; private set; } = "A4";
        public string ExportFormat { get; private set; } = "PDF";
        public bool EnableWordExport { get; private set; } = false;

        public SettingsForm(MainForm mainForm, bool skipZeroTotal, string paperSize, string exportFormat)
        {
            InitializeComponent();

            _mainForm = mainForm;

            if (chkSkipZeroTotal != null)
                chkSkipZeroTotal.Checked = Properties.Settings.Default.SkipZeroTotal;

            EnableWordExport = Properties.Settings.Default.EnableWordExport;

            var savedFormat = Properties.Settings.Default.ExportFormat;
            var formatIndex = cmbExportFormat.Items.IndexOf(savedFormat);
            cmbExportFormat.SelectedIndex = formatIndex >= 0 ? formatIndex : 0;   // default PDF
            cmbExportFormat.Enabled = EnableWordExport;

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

            var cachePath = _mainForm.GetAccountsCachePath();
            linkLabelClearCache.Enabled = File.Exists(cachePath);

            //UpdateCacheStatus();
        }

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

        private void linkLabelClearCache_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
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
                var path = _mainForm.GetAccountsCachePath();
                if (File.Exists(path))
                    File.Delete(path);

                _mainForm?.ClearCachedRecords();

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
    }
}
