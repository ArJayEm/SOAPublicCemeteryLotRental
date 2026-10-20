using System.Diagnostics;
using System.Reflection;

namespace StatementOfAccountPrinting
{
    public partial class About : Form
    {
        public About()
        {
            InitializeComponent();

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
            lblVersion.Text = $"v{cleanVersion}";
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var text = linkLabel1.Text; // or hardcode if you prefer

            if (string.IsNullOrWhiteSpace(text))
                return;

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = text,
                UseShellExecute = true   // ← required to open the default browser
            });
        }
    }
}
