namespace StatementOfAccountPrinting
{
    public partial class WhatsNewForm : Form
    {
        private string _url;

        public WhatsNewForm(string url)
        {
            InitializeComponent();

            _url = url;
            Text = "What's New?";
            Size = new Size(850, 650);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;

            // Load the URL when the form opens
            this.Load += async (s, e) => await InitializeBrowser();
        }

        private async Task InitializeBrowser()
        {
            try
            {
                await webView21.EnsureCoreWebView2Async(null);

                // Download the HTML as text
                using var client = new System.Net.Http.HttpClient();
                string htmlContent = await client.GetStringAsync(_url);

                // Render it manually
                webView21.NavigateToString(htmlContent);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not load changelog: " + ex.Message);
            }
        }

        private void WhatsNewForm_Load(object sender, EventArgs e)
        {

        }
    }
}
