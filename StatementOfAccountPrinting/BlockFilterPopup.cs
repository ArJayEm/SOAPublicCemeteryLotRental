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

        public void LoadBlocks(IEnumerable<string> blocks, IEnumerable<string>? checkedBlocks = null)
        {
            _loading = true;
            try
            {
                var checkedSet = checkedBlocks == null
                    ? null
                    : new HashSet<string>(checkedBlocks, StringComparer.OrdinalIgnoreCase);

                clbBlockFilter.Items.Clear();
                foreach (var b in blocks)
                {
                    // Rules:
                    //   checkedBlocks == null          → check all (default state)
                    //   checkedBlocks has items        → check only those
                    //   checkedBlocks is empty set     → check nothing
                    bool isChecked;
                    if (checkedSet == null)
                        isChecked = true;
                    else
                        isChecked = checkedSet.Contains(b);

                    clbBlockFilter.Items.Add(b, isChecked);
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
                if (item != null) result.Add(item.ToString()!);
            return result;
        }

        public void ShowUnder(Control anchor)
        {
            var pt = anchor.PointToScreen(new System.Drawing.Point(0, anchor.Height));
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