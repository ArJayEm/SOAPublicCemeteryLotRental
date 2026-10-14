using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

public static class SoaTemplateReader
{
    public const string TemplateFileName = "soa-public_cemetery_lot_rental.docx";
    public const string TemplateFolder = "Templates";

    /// <summary>Locates the template file (same logic used by WordTemplateService).</summary>
    public static string ResolveTemplatePath()
    {
        var exePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                   TemplateFolder, TemplateFileName);
        if (File.Exists(exePath)) return exePath;

        var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
        while (dir != null)
        {
            var candidate = Path.Combine(dir.FullName, TemplateFolder, TemplateFileName);
            if (File.Exists(candidate)) return candidate;
            dir = dir.Parent;
        }

        var cwdPath = Path.Combine(Directory.GetCurrentDirectory(),
                                   TemplateFolder, TemplateFileName);
        if (File.Exists(cwdPath)) return cwdPath;

        throw new FileNotFoundException(
            $"Template '{TemplateFileName}' not found.", TemplateFileName);
    }

    /// <summary>
    /// Reads every paragraph in every header part of the .docx, returning
    /// them as a list of plain text lines (blank lines removed).
    /// </summary>
    public static List<string> ReadHeaders(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, false);
        var result = new List<string>();

        foreach (var headerPart in doc.MainDocumentPart!.HeaderParts)
        {
            var paras = headerPart.Header.Descendants<Paragraph>()
                .Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text)))
                .Where(s => !string.IsNullOrWhiteSpace(s));

            result.AddRange(paras);
        }

        return result;
    }

    /// <summary>Same as ReadHeaders, but for footers.</summary>
    public static List<string> ReadFooters(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, false);
        var result = new List<string>();

        foreach (var footerPart in doc.MainDocumentPart.FooterParts)
        {
            var paras = footerPart.Footer.Descendants<Paragraph>()
                .Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text)))
                .Where(s => !string.IsNullOrWhiteSpace(s));

            result.AddRange(paras);
        }

        return result;
    }

    /// <summary>Extracts embedded images from header/footer parts (for logo support).</summary>
    public static List<byte[]> ReadHeaderImages(string docxPath)
    {
        var images = new List<byte[]>();

        using var doc = WordprocessingDocument.Open(docxPath, false);
        foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
        {
            foreach (var blip in headerPart.Header.Descendants<DocumentFormat.OpenXml.Drawing.Blip>())
            {
                var relId = blip.Embed?.Value;
                if (string.IsNullOrEmpty(relId)) continue;

                var imagePart = headerPart.GetPartById(relId) as ImagePart;
                if (imagePart == null) continue;

                using var stream = imagePart.GetStream();
                using var ms = new MemoryStream();
                stream.CopyTo(ms);
                images.Add(ms.ToArray());
            }
        }

        return images;
    }

    public static List<byte[]> ReadFooterImages(string docxPath)
    {
        var images = new List<byte[]>();

        using var doc = WordprocessingDocument.Open(docxPath, false);
        foreach (var footerPart in doc.MainDocumentPart.FooterParts)
        {
            foreach (var blip in footerPart.Footer.Descendants<DocumentFormat.OpenXml.Drawing.Blip>())
            {
                var relId = blip.Embed?.Value;
                if (string.IsNullOrEmpty(relId)) continue;

                var imagePart = footerPart.GetPartById(relId) as ImagePart;
                if (imagePart == null) continue;

                using var stream = imagePart.GetStream();
                using var ms = new MemoryStream();
                stream.CopyTo(ms);
                images.Add(ms.ToArray());
            }
        }

        return images;
    }
}