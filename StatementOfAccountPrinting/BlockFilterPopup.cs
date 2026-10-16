namespace StatementOfAccountPrinting
{
    public partial class BlockFilterPopup : Form
    {
        public event Action<HashSet<string>>? SelectionChanged;

        private bool _loading;

        public BlockFilterPopup()
        {
            InitializeComponent();

            // Close when focus leaves the popup
            Deactivate += (s, e) => Hide();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            // Wire the event handler only after the handle exists.
            // Avoid double-subscribe by removing first.
            clbBlockFilter.ItemCheck -= clbBlockFilter_ItemCheck;
            clbBlockFilter.ItemCheck += clbBlockFilter_ItemCheck;
        }

        public void LoadBlocks(IEnumerable<(string Name, int Count)> blocks,
                       IEnumerable<string>? checkedBlocks = null)
        {
            _loading = true;
            try
            {
                var checkedSet = checkedBlocks == null
                    ? null
                    : new HashSet<string>(checkedBlocks, StringComparer.OrdinalIgnoreCase);

                clbBlockFilter.Items.Clear();
                foreach (var (name, count) in blocks)
                {
                    bool isChecked = checkedSet == null || checkedSet.Contains(name);
                    clbBlockFilter.Items.Add($"{name} ({count})", isChecked);
                }
            }
            finally
            {
                _loading = false;
            }
        }

        public HashSet<string> GetCheckedBlocks()
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in clbBlockFilter.CheckedItems)
            {
                var s = item?.ToString() ?? "";
                if (string.IsNullOrEmpty(s)) continue;

                // Remove the trailing " (N)" if present
                int paren = s.LastIndexOf(" (");
                if (paren > 0 && s.EndsWith(")"))
                    s = s.Substring(0, paren);

                result.Add(s);
            }
            return result;
        }

        public void ShowUnder(ToolStripItem anchor)
        {
            if (anchor?.Owner == null) return;

            var strip = anchor.Owner;   // the ToolStrip control
            var pt = strip.PointToScreen(new System.Drawing.Point(anchor.Bounds.Left,
                                                                     anchor.Bounds.Bottom));
            Location = pt;
            Show();
            BringToFront();
            clbBlockFilter.Focus();
        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            _loading = true;
            try
            {
                for (int i = 0; i < clbBlockFilter.Items.Count; i++)
                    clbBlockFilter.SetItemChecked(i, true);
            }
            finally
            {
                _loading = false;
            }
            SelectionChanged?.Invoke(GetCheckedBlocks());
        }

        private void btnSelectNone_Click(object sender, EventArgs e)
        {
            _loading = true;
            try
            {
                for (int i = 0; i < clbBlockFilter.Items.Count; i++)
                    clbBlockFilter.SetItemChecked(i, false);
            }
            finally
            {
                _loading = false;
            }
            SelectionChanged?.Invoke(GetCheckedBlocks());
        }

        private void clbBlockFilter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            // Suppress while we're bulk-loading items or checking them programmatically
            if (_loading) return;

            if (!IsHandleCreated)
            {
                SelectionChanged?.Invoke(GetCheckedBlocks());
                return;
            }

            // Fire after the check state actually changes
            BeginInvoke(new Action(() =>
                SelectionChanged?.Invoke(GetCheckedBlocks())));
        }
    }
}