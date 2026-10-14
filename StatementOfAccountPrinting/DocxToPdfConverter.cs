using System;
using System.IO;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

public static class DocxToPdfConverter
{
    /// <summary>
    /// Converts a .docx to .pdf using Microsoft Word COM automation.
    /// Requires Word to be installed.
    /// </summary>
    public static void Convert(string docxPath, string pdfPath)
    {
        if (!File.Exists(docxPath))
            throw new FileNotFoundException("Source .docx not found.", docxPath);

        var fullDocx = Path.GetFullPath(docxPath);
        var fullPdf = Path.GetFullPath(pdfPath);

        // Ensure output directory exists
        Directory.CreateDirectory(Path.GetDirectoryName(fullPdf)!);

        Word.Application wordApp = null;
        Word.Document document = null;

        try
        {
            wordApp = new Word.Application
            {
                Visible = false,
                ScreenUpdating = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
            };

            // Open the doc — read-only is fine, we're only exporting
            document = wordApp.Documents.Open(
                FileName: fullDocx,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false);

            // 17 = wdExportFormatPDF
            document.ExportAsFixedFormat(
                OutputFileName: fullPdf,
                ExportFormat: Word.WdExportFormat.wdExportFormatPDF,
                OpenAfterExport: false,
                OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                Range: Word.WdExportRange.wdExportAllDocument,
                Item: Word.WdExportItem.wdExportDocumentWithMarkup,
                IncludeDocProps: true,
                KeepIRM: true,
                CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks,
                DocStructureTags: true,
                BitmapMissingFonts: true,
                UseISO19005_1: false);
        }
        finally
        {
            // ── Always release COM objects ──
            if (document != null)
            {
                document.Close(SaveChanges: false);
                Marshal.ReleaseComObject(document);
                document = null;
            }

            if (wordApp != null)
            {
                wordApp.Quit(SaveChanges: false);
                Marshal.ReleaseComObject(wordApp);
                wordApp = null;
            }

            // Force garbage collection so Word fully exits
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
    public static bool IsWordInstalled()
    {
        try
        {
            var type = Type.GetTypeFromProgID("Word.Application");
            return type != null;
        }
        catch
        {
            return false;
        }
    }
}