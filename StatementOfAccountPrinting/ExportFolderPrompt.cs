namespace StatementOfAccountPrinting
{
    public partial class ExportFolderPrompt : Form
    {
        public ExportFolderPrompt(string suggestedFolder)
        {
            Text = "Export Folder";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ClientSize = new Size(450, 140);
            BackColor = Color.White;

            var lblMessage = new Label
            {
                Text = "A folder with the same name as your imported file will be created (or was found):\n\n" +
                       suggestedFolder + "\\\n\n" +
                       "Do you want to use this folder?",
                Left = 15,
                Top = 15,
                Width = 420,
                Height = 80
            };

            var btnYes = new Button
            {
                Text = "Yes, use this folder",
                Left = 15,
                Top = 100,
                Width = 140,
                DialogResult = DialogResult.Yes
            };

            var btnCustom = new Button
            {
                Text = "Choose my own folder",
                Left = 165,
                Top = 100,
                Width = 140,
                DialogResult = DialogResult.No
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                Left = 315,
                Top = 100,
                Width = 80,
                DialogResult = DialogResult.Cancel
            };

            Controls.Add(lblMessage);
            Controls.Add(btnYes);
            Controls.Add(btnCustom);
            Controls.Add(btnCancel);

            AcceptButton = btnYes;
            CancelButton = btnCancel;
        }
    }
}
