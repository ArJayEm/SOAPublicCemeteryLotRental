public class ProgressDialog : Form
{
    private ProgressBar _bar;
    private Label _label;
    private Label _percent;
    private Button _close;

    public ProgressDialog(string title, string initialMessage = "Working…")
    {
        Text = title;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;
        MaximizeBox = false;
        ControlBox = false;
        ClientSize = new System.Drawing.Size(460, 240);

        _label = new Label
        {
            Text = initialMessage,
            Left = 12,
            Top = 12,
            Width = 436,
            Height = 110,
            AutoSize = false,
            AutoEllipsis = false
        };

        _bar = new ProgressBar
        {
            Left = 12,
            Top = 135,
            Width = 436,
            Height = 20,
            Minimum = 0,
            Maximum = 100,
            Style = ProgressBarStyle.Continuous
        };

        _percent = new Label
        {
            Text = "0%",
            Left = 12,
            Top = 170,
            Width = 260,
            Height = 24,
            TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        };

        _close = new Button
        {
            Text = "Close",
            Left = 350,
            Top = 170,
            Width = 98,
            Height = 26,
            Visible = false
        };
        _close.Click += (s, e) => Close();

        Controls.Add(_label);
        Controls.Add(_bar);
        Controls.Add(_percent);
        Controls.Add(_close);
    }

    public void Report(int percent, string message)
    {
        if (IsDisposed) return;

        if (InvokeRequired)
        {
            BeginInvoke(new Action(() => Report(percent, message)));
            return;
        }

        _bar.Value = Math.Max(_bar.Minimum, Math.Min(_bar.Maximum, percent));
        _percent.Text = $"{_bar.Value}%";
        if (!string.IsNullOrEmpty(message))
            _label.Text = message;
    }

    public void ShowCloseButton()
    {
        if (IsDisposed) return;

        if (InvokeRequired)
        {
            BeginInvoke(new Action(ShowCloseButton));
            return;
        }

        _close.Visible = true;
        _close.Focus();
    }

    /// <summary>
    /// Runs the work; if keepOpenWhenDone is true, the dialog stays open
    /// until the user clicks Close. Only then does RunAsync return.
    /// </summary>
    public static async Task<bool> RunAsync(
    IWin32Window owner,
    string title,
    Func<IProgress<(int Pct, string Message)>, Task> work,
    bool keepOpenWhenDone = true)
    {
        using var dlg = new ProgressDialog(title);
        var progress = new Progress<(int Pct, string Message)>(p => dlg.Report(p.Pct, p.Message));
        bool ok = false;

        dlg.Shown += async (s, e) =>
        {
            try
            {
                await work(progress);
                ok = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner,
                    "Operation failed:\n\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (keepOpenWhenDone) dlg.ShowCloseButton();
                else dlg.Close();
            }
        };

        dlg.ShowDialog(owner);
        return ok;
    }
}