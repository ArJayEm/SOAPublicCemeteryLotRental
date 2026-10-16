namespace StatementOfAccountPrinting
{
    public partial class About : Form
    {
        public About()
        {
            InitializeComponent();

            var version = System.Reflection.Assembly
                .GetExecutingAssembly()
                .GetName()
                .Version?
                .ToString() ?? "unknown";
            lblVersion.Text = $"Version: {version}";
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
