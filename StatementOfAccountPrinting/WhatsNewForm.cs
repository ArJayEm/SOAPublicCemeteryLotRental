using System.Text.Json;

namespace StatementOfAccountPrinting
{
    public partial class WhatsNewForm : Form
    {
        private string _jsonUrl;
        private string _currentVersion;
        private FlowLayoutPanel _mainPanel;

        public WhatsNewForm(string jsonUrl, string currentVersion)
        {
            _jsonUrl = jsonUrl;
            _currentVersion = currentVersion;

            Text = "What's New?";
            Size = new Size(650, 600);
            MinimumSize = new Size(500, 400); // 👈 Add a minimum size
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.FromArgb(243, 244, 246);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;

            _mainPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(15),
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
            };
            Controls.Add(_mainPanel);

            this.Load += async (s, e) => await LoadChangelog();
        }

        private async Task LoadChangelog()
        {
            _mainPanel.Controls.Clear();
            var loadingLabel = new Label { Text = "Loading changelog...", AutoSize = true, Font = new Font("Segoe UI", 12) };
            _mainPanel.Controls.Add(loadingLabel);

            // Wait for the form to render to get an accurate width
            await Task.Delay(200);

            try
            {
                using var client = new HttpClient();
                client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0");
                string json = await client.GetStringAsync(_jsonUrl);

                var releases = JsonSerializer.Deserialize<List<ChangelogRelease>>(json);

                _mainPanel.Controls.Clear();

                if (releases == null || releases.Count == 0) return;

                bool isFirst = true;
                foreach (var release in releases)
                {
                    var card = BuildReleaseCard(release, isFirst);
                    _mainPanel.Controls.Add(card);
                    isFirst = false;
                }
            }
            catch (Exception ex)
            {
                _mainPanel.Controls.Clear();
                _mainPanel.Controls.Add(new Label { Text = "Error: " + ex.Message, ForeColor = Color.Red, AutoSize = true });
            }
        }

        private Panel BuildReleaseCard(ChangelogRelease release, bool isExpanded)
        {
            // 1. Create the main card (TableLayoutPanel keeps layout perfect)
            var card = new TableLayoutPanel
            {
                Width = _mainPanel.ClientSize.Width > 0 ? _mainPanel.ClientSize.Width - 40 : 550,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                Margin = new Padding(0, 0, 0, 15),
                ColumnCount = 1,
                RowCount = 2
            };
            card.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            card.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));
            card.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // 2. Create Header and its labels
            var header = new Panel { Dock = DockStyle.Fill, Cursor = Cursors.Hand, BackColor = Color.White };

            var arrow = new Label { Text = isExpanded ? "▼" : "▶", Font = new Font("Segoe UI", 9), AutoSize = true, Location = new Point(10, 15), Cursor = Cursors.Hand };
            var lblVersion = new Label { Text = release.Version, Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 215), Location = new Point(30, 13), Cursor = Cursors.Hand, Width = card.Width };
            var lblDate = new Label { Text = release.Date, Font = new Font("Segoe UI", 9), ForeColor = Color.Gray, AutoSize = true, Anchor = AnchorStyles.Top | AnchorStyles.Right, Cursor = Cursors.Hand };

            header.Controls.Add(arrow);
            header.Controls.Add(lblVersion);
            header.Controls.Add(lblDate);

            // 3. Create Content panel (Must be created BEFORE the click events)
            var content = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(30, 5, 20, 15),
                Visible = isExpanded
            };

            // 4. Add sections to Content
            if (release.Sections != null)
            {
                foreach (var section in release.Sections)
                {
                    var badge = new Label
                    {
                        Text = section.Title?.ToUpper() ?? "UPDATES",
                        Font = new Font("Segoe UI", 8, FontStyle.Bold),
                        ForeColor = Color.White,
                        BackColor = GetBadgeColor(section.Type),
                        AutoSize = true,
                        Padding = new Padding(5, 2, 5, 2),
                        Margin = new Padding(0, 10, 0, 5)
                    };

                    var list = new Label
                    {
                        Text = "• " + string.Join("\n• ", section.Items ?? new List<string>()),
                        Font = new Font("Segoe UI", 9),
                        AutoSize = true,
                        MaximumSize = new Size(card.Width - 80, 0),
                        Margin = new Padding(0, 0, 0, 5)
                    };

                    content.Controls.Add(badge);
                    content.Controls.Add(list);
                }
            }

            // 5. Position Date Label
            header.Resize += (s, e) => lblDate.Location = new Point(header.Width - lblDate.Width - 15, 15);
            lblDate.Location = new Point(header.Width - lblDate.Width - 15, 15);

            // 6. Setup Click Events (Now that 'content' is safely declared)
            Action toggleCard = () =>
            {
                content.Visible = !content.Visible;
                arrow.Text = content.Visible ? "▼" : "▶";
            };

            header.Click += (s, e) => toggleCard();
            arrow.Click += (s, e) => toggleCard();
            lblVersion.Click += (s, e) => toggleCard();
            lblDate.Click += (s, e) => toggleCard();

            // 7. Add controls to the card
            card.Controls.Add(header, 0, 0);
            card.Controls.Add(content, 0, 1);

            return card;
        }

        private Color GetBadgeColor(string type)
        {
            if (string.IsNullOrEmpty(type)) return Color.Gray;

            return type.ToLower() switch
            {
                "update" => Color.FromArgb(16, 124, 16),
                "fix" => Color.FromArgb(209, 52, 56),
                "feature" => Color.FromArgb(0, 90, 158),
                "improvement" => Color.FromArgb(135, 100, 184),
                _ => Color.Gray
            };
        }
    }
}