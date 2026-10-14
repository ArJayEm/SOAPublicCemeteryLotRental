using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StatementOfAccountPrinting
{
    public partial class SettingsForm : Form
    {
        public event Action? CacheCleared;

        public SettingsForm()
        {
            InitializeComponent();

            cmbPaperSize.Items.Clear();
            cmbPaperSize.Items.Add("A4");
            cmbPaperSize.Items.Add("Legal");
            cmbPaperSize.SelectedIndex = 0;   // default to A4

            // ── Load saved settings ──
            if (chkSkipZeroTotal != null)
                chkSkipZeroTotal.Checked = Properties.Settings.Default.SkipZeroTotal;

            if (cmbPaperSize != null)
            {
                var saved = Properties.Settings.Default.PaperSize;
                var idx = cmbPaperSize.Items.IndexOf(saved);
                cmbPaperSize.SelectedIndex = idx >= 0 ? idx : 0;
            }

            UpdateCacheStatus();
        }

        private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Save();
        }

        private void Save()
        {
            try
            {
                if (chkSkipZeroTotal != null)
                    Properties.Settings.Default.SkipZeroTotal = chkSkipZeroTotal.Checked;

                if (cmbPaperSize?.SelectedItem != null)
                    Properties.Settings.Default.PaperSize = cmbPaperSize.SelectedItem.ToString();

                Properties.Settings.Default.Save();
            }
            catch { }
        }

        private CheckBox _chkSkipZero;
        private ComboBox _cmbPaperSize;
        private Button _btnOk;
        private Button _btnCancel;

        public bool SkipZeroTotal { get; private set; }
        public string PaperSize { get; private set; } = "A4";

        public SettingsForm(bool skipZeroTotal, string paperSize)
        {
            // ── Form basics ──
            Text = "Settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ClientSize = new System.Drawing.Size(360, 160);

            // ── Skip zero totals ──
            _chkSkipZero = new CheckBox
            {
                Text = "Skip records with ₱0.00 total on export",
                Left = 12,
                Top = 16,
                Width = 330,
                Checked = skipZeroTotal
            };

            // ── Paper size ──
            var lblPaper = new Label
            {
                Text = "PDF paper size:",
                Left = 12,
                Top = 56,
                Width = 120
            };

            _cmbPaperSize = new ComboBox
            {
                Left = 140,
                Top = 52,
                Width = 120,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbPaperSize.Items.AddRange(new object[] { "A4", "Legal" });
            _cmbPaperSize.SelectedItem = _cmbPaperSize.Items.Contains(paperSize) ? paperSize : "A4";

            // ── Buttons ──
            _btnOk = new Button
            {
                Text = "OK",
                Left = 170,
                Top = 110,
                Width = 80,
                DialogResult = DialogResult.OK
            };

            _btnCancel = new Button
            {
                Text = "Cancel",
                Left = 260,
                Top = 110,
                Width = 80,
                DialogResult = DialogResult.Cancel
            };

            // Accept Enter/Escape
            AcceptButton = _btnOk;
            CancelButton = _btnCancel;

            // Commit values when OK is clicked
            _btnOk.Click += (s, e) =>
            {
                SkipZeroTotal = _chkSkipZero.Checked;
                PaperSize = _cmbPaperSize.SelectedItem?.ToString() ?? "A4";
            };

            Controls.Add(_chkSkipZero);
            Controls.Add(lblPaper);
            Controls.Add(_cmbPaperSize);
            Controls.Add(_btnOk);
            Controls.Add(_btnCancel);
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

        private void UpdateCacheStatus()
        {
            var path = GetAccountsCachePath();
            var message = File.Exists(path)
                ? $"Cache file: {path}"
                : "No cache file present.";

            MessageBox.Show(message);
        }

        private static string GetAccountsCachePath()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "StatementOfAccountPrinting");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "accounts.json");
        }
        #endregion
    }
}
