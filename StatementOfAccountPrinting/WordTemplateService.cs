using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using StatementOfAccountPrinting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// ✅ Alias only the Wordprocessing Text class to avoid ambiguity
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;

public static class WordTemplateService
{
    public const string TemplateFileName = "soa-public_cemetery_lot_rental.docx";
    public const string TemplateFolder = "Templates";

    // ═════════════════════════════════════════════════════════════
    // RESOLVE TEMPLATE PATH
    // ═════════════════════════════════════════════════════════════
    public static string ResolveTemplatePath()
    {
        var exePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            TemplateFolder,
            TemplateFileName);

        if (File.Exists(exePath)) return exePath;

        var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
        while (dir != null)
        {
            var candidate = Path.Combine(dir.FullName, TemplateFolder, TemplateFileName);
            if (File.Exists(candidate)) return candidate;
            dir = dir.Parent;
        }

        var cwdPath = Path.Combine(
            Directory.GetCurrentDirectory(),
            TemplateFolder,
            TemplateFileName);

        if (File.Exists(cwdPath)) return cwdPath;

        throw new FileNotFoundException(
            $"Template '{TemplateFileName}' not found. Looked in:\n" +
            $"  • {exePath}\n" +
            $"  • {cwdPath}\n\n" +
            "Make sure the file is included in the project with " +
            "Build Action = Content and Copy to Output Directory = Copy if newer.",
            TemplateFileName);
    }

    // ═════════════════════════════════════════════════════════════
    // GENERATE WORD (fill template with account)
    // ═════════════════════════════════════════════════════════════
    public static void GenerateWord(string outputPath, CemeteryLotAccount account)
    {
        if (account == null) throw new ArgumentNullException(nameof(account));

        var templatePath = ResolveTemplatePath();

        if (File.Exists(outputPath)) File.Delete(outputPath);
        File.Copy(templatePath, outputPath, overwrite: true);

        using var doc = WordprocessingDocument.Open(outputPath, isEditable: true);
        var main = doc.MainDocumentPart!;
        var body = main.Document.Body!;
        var map = ToPlaceholders(account);

        // Body
        foreach (var para in body.Descendants<Paragraph>())
        {
            var paraText = string.Concat(para.Descendants<WText>().Select(t => t.Text));
            if (paraText.Contains("{{"))
            {
                FlattenRuns(para);
                ReplaceTokens(para, map);
            }
        }

        // Headers
        foreach (var hp in main.HeaderParts)
        {
            foreach (var para in hp.Header.Descendants<Paragraph>())
            {
                var t = string.Concat(para.Descendants<WText>().Select(x => x.Text));
                if (t.Contains("{{"))
                {
                    FlattenRuns(para);
                    ReplaceTokens(para, map);
                }
            }
            hp.Header.Save();
        }

        // Footers
        foreach (var fp in main.FooterParts)
        {
            foreach (var para in fp.Footer.Descendants<Paragraph>())
            {
                var t = string.Concat(para.Descendants<WText>().Select(x => x.Text));
                if (t.Contains("{{"))
                {
                    FlattenRuns(para);
                    ReplaceTokens(para, map);
                }
            }
            fp.Footer.Save();
        }

        main.Document.Save();
    }

    // ═════════════════════════════════════════════════════════════
    // FILL TEMPLATE (template path + account + output path)
    // ═════════════════════════════════════════════════════════════
    public static void FillTemplate(string templatePath, CemeteryLotAccount account, string outputPath)
    {
        if (account == null) throw new ArgumentNullException(nameof(account));

        if (File.Exists(outputPath)) File.Delete(outputPath);
        File.Copy(templatePath, outputPath, overwrite: true);

        using var doc = WordprocessingDocument.Open(outputPath, isEditable: true);
        var main = doc.MainDocumentPart!;
        var body = main.Document.Body!;
        var map = ToPlaceholders(account);

        // Body
        foreach (var para in body.Descendants<Paragraph>())
        {
            var paraText = string.Concat(para.Descendants<WText>().Select(t => t.Text));
            if (paraText.Contains("{{"))
            {
                FlattenRuns(para);
                ReplaceTokens(para, map);
            }
        }

        // Headers
        foreach (var hp in main.HeaderParts)
        {
            foreach (var para in hp.Header.Descendants<Paragraph>())
            {
                var t = string.Concat(para.Descendants<WText>().Select(x => x.Text));
                if (t.Contains("{{"))
                {
                    FlattenRuns(para);
                    ReplaceTokens(para, map);
                }
            }
            hp.Header.Save();
        }

        // Footers
        foreach (var fp in main.FooterParts)
        {
            foreach (var para in fp.Footer.Descendants<Paragraph>())
            {
                var t = string.Concat(para.Descendants<WText>().Select(x => x.Text));
                if (t.Contains("{{"))
                {
                    FlattenRuns(para);
                    ReplaceTokens(para, map);
                }
            }
            fp.Footer.Save();
        }

        main.Document.Save();
    }

    // ═════════════════════════════════════════════════════════════
    // CREATE TEMPLATE (blank copy)
    // ═════════════════════════════════════════════════════════════
    public static void CreateTemplate(string outputPath)
    {
        var templatePath = ResolveTemplatePath();

        if (!File.Exists(templatePath))
            throw new FileNotFoundException(
                $"Template file not found: {templatePath}", templatePath);

        if (File.Exists(outputPath))
            File.Delete(outputPath);

        File.Copy(templatePath, outputPath, overwrite: true);
    }

    // ═════════════════════════════════════════════════════════════
    // CREATE TEMPLATE (copy + fill with account)
    // ═════════════════════════════════════════════════════════════
    public static void CreateTemplate(string outputPath, CemeteryLotAccount account)
    {
        if (account == null) throw new ArgumentNullException(nameof(account));

        CreateTemplate(outputPath);

        using var doc = WordprocessingDocument.Open(outputPath, isEditable: true);
        var main = doc.MainDocumentPart!;
        var body = main.Document.Body!;
        var map = ToPlaceholders(account);

        // Body
        foreach (var para in body.Descendants<Paragraph>())
        {
            var paraText = string.Concat(para.Descendants<WText>().Select(t => t.Text));
            if (paraText.Contains("{{"))
            {
                FlattenRuns(para);
                ReplaceTokens(para, map);
            }
        }

        // Headers
        foreach (var hp in main.HeaderParts)
        {
            foreach (var para in hp.Header.Descendants<Paragraph>())
            {
                var t = string.Concat(para.Descendants<WText>().Select(x => x.Text));
                if (t.Contains("{{"))
                {
                    FlattenRuns(para);
                    ReplaceTokens(para, map);
                }
            }
            hp.Header.Save();
        }

        // Footers
        foreach (var fp in main.FooterParts)
        {
            foreach (var para in fp.Footer.Descendants<Paragraph>())
            {
                var t = string.Concat(para.Descendants<WText>().Select(x => x.Text));
                if (t.Contains("{{"))
                {
                    FlattenRuns(para);
                    ReplaceTokens(para, map);
                }
            }
            fp.Footer.Save();
        }

        main.Document.Save();
    }

    // ═════════════════════════════════════════════════════════════
    // PLACEHOLDER MAPPING
    // ═════════════════════════════════════════════════════════════
    public static Dictionary<string, string> ToPlaceholders(CemeteryLotAccount a)
    {
        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["DateGenerated"] = DateTime.Today.ToString("MMMM dd, yyyy"),
            ["NAME OF TAXPAYERS"] = (a.LesseeName ?? "").Trim(),
            ["ADDRESS"] = (a.Address ?? "").Trim(),
            ["LotBlockNumber"] = (a.LotBlockNumber ?? "").Trim(),
            ["LOT AREA"] = (a.AreaSqm ?? "").Trim(),
            ["NAME OF DECEASED"] = (a.NameOfDeceased ?? "").Trim(),
            ["TOTAL BASIC TAX-Final"] = a.RentalOrRenewalFee ?? "0.00",
            ["SURCHARGE"] = a.Surcharge ?? "0.00",
            ["INTEREST"] = a.Interest ?? "0.00",
            ["TOTAL DELIQUENCY"] = a.TotalAmountDue ?? "0.00",
        };
    }

    // ═════════════════════════════════════════════════════════════
    // FLATTEN RUNS
    // ═════════════════════════════════════════════════════════════
    private static void FlattenRuns(Paragraph para)
    {
        if (para == null) return;

        var runs = para.Elements<Run>().ToList();
        if (runs.Count <= 1) return;

        var combined = string.Concat(
            runs.SelectMany(r => r.Elements<WText>()).Select(t => t.Text ?? string.Empty));

        var firstRun = runs[0];

        foreach (var t in firstRun.Elements<WText>().ToList())
            t.Remove();

        firstRun.AppendChild(new WText(combined)
        {
            Space = SpaceProcessingModeValues.Preserve
        });

        foreach (var r in runs.Skip(1))
            r.Remove();
    }

    // ═════════════════════════════════════════════════════════════
    // REPLACE TOKENS (+ underline value only)
    // ═════════════════════════════════════════════════════════════
    private static readonly HashSet<string> UnderlineKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "DateGenerated",
        "NAME OF TAXPAYERS",
        "ADDRESS",
        "LotBlockNumber",
        "LOT AREA",
        "NAME OF DECEASED",
    };

    private static void ReplaceTokens(Paragraph para, Dictionary<string, string> map)
    {
        if (para == null || map == null) return;

        foreach (var text in para.Descendants<WText>().ToList())
        {
            var original = text.Text ?? "";
            var replaced = original;
            string? underlineKey = null;

            foreach (var kvp in map)
            {
                var token = "{{" + kvp.Key + "}}";
                if (replaced.IndexOf(token, StringComparison.Ordinal) >= 0)
                {
                    replaced = replaced.Replace(token, kvp.Value ?? "");
                    if (UnderlineKeys.Contains(kvp.Key))
                        underlineKey = kvp.Key;
                }
            }

            if (string.Equals(original, replaced, StringComparison.Ordinal))
                continue;

            // Handle any embedded newlines in the replacement
            if (replaced.Contains('\n'))
                InsertLineBreaks(text, replaced);
            else
                text.Text = replaced;

            text.Space = SpaceProcessingModeValues.Preserve;

            // Split the run so only the substituted value is underlined
            if (underlineKey != null)
                UnderlineValueOnly(text, map[underlineKey]);
        }
    }

    // ═════════════════════════════════════════════════════════════
    // UNDERLINE VALUE ONLY (label stays plain)
    // ═════════════════════════════════════════════════════════════
    private static void UnderlineValueOnly(WText text, string value)
    {
        if (text == null || string.IsNullOrEmpty(value)) return;

        var run = text.Parent as Run;
        if (run == null) return;

        var full = text.Text ?? "";
        int idx = full.IndexOf(value, StringComparison.Ordinal);
        if (idx < 0) return;

        var before = full.Substring(0, idx);
        var after = full.Substring(idx + value.Length);

        RunProperties? originalProps = run.RunProperties;

        // before
        if (before.Length > 0)
        {
            var beforeRun = new Run();
            if (originalProps != null)
                beforeRun.AppendChild(originalProps.CloneNode(true));
            beforeRun.AppendChild(new WText(before) { Space = SpaceProcessingModeValues.Preserve });
            run.InsertBeforeSelf(beforeRun);
        }

        // value (underlined)
        var valueRun = new Run();
        RunProperties valueProps = originalProps != null
            ? (RunProperties)originalProps.CloneNode(true)
            : new RunProperties();

        foreach (var u in valueProps.Elements<Underline>().ToList())
            u.Remove();
        valueProps.Append(new Underline { Val = UnderlineValues.Single });

        valueRun.AppendChild(valueProps);
        valueRun.AppendChild(new WText(value) { Space = SpaceProcessingModeValues.Preserve });
        run.InsertBeforeSelf(valueRun);

        // after
        if (after.Length > 0)
        {
            var afterRun = new Run();
            if (originalProps != null)
                afterRun.AppendChild(originalProps.CloneNode(true));
            afterRun.AppendChild(new WText(after) { Space = SpaceProcessingModeValues.Preserve });
            run.InsertBeforeSelf(afterRun);
        }

        // remove original
        text.Remove();
        if (!run.Elements<WText>().Any() &&
            !run.Elements<Break>().Any() &&
            !run.Elements<Drawing>().Any())
        {
            run.Remove();
        }
    }
    /// <summary>
    /// Generates ONE Word document containing all accounts, one page per account,
    /// using the bundled template as a base. Each account gets a copy of the
    /// template body with placeholders replaced, separated by a page break.
    /// </summary>
    public static void GenerateCombinedWord(string outputPath, List<CemeteryLotAccount> accounts)
    {
        if (accounts == null || accounts.Count == 0)
            throw new ArgumentException("No accounts to generate.", nameof(accounts));

        var templatePath = ResolveTemplatePath();

        if (File.Exists(outputPath)) File.Delete(outputPath);
        File.Copy(templatePath, outputPath, overwrite: true);

        using var doc = WordprocessingDocument.Open(outputPath, isEditable: true);
        var main = doc.MainDocumentPart!;
        var body = main.Document.Body!;

        // 1) Capture the original body content (the blank template body)
        var originalChildren = body.ChildElements
            .Select(e => e.CloneNode(true))
            .ToList();

        // 2) Clear the body — we'll re-add filled copies per account
        body.RemoveAllChildren();

        // 3) For each account, append a filled clone of the template body
        for (int i = 0; i < accounts.Count; i++)
        {
            var account = accounts[i];
            var map = ToPlaceholders(account);

            foreach (var node in originalChildren)
            {
                var clone = node.CloneNode(true);

                // Replace tokens inside every cloned paragraph
                foreach (var para in clone.Descendants<Paragraph>())
                {
                    var paraText = string.Concat(
                        para.Descendants<WText>().Select(t => t.Text));
                    if (paraText.Contains("{{"))
                    {
                        FlattenRuns(para);
                        ReplaceTokens(para, map);
                    }
                }

                body.AppendChild(clone);
            }

            // 4) Insert a page break between accounts (but not after the last one)
            if (i < accounts.Count - 1)
            {
                body.AppendChild(new Paragraph(
                    new Run(new Break { Type = BreakValues.Page })));
            }
        }

        main.Document.Save();
    }

    // ═════════════════════════════════════════════════════════════
    // INSERT LINE BREAKS
    // ═════════════════════════════════════════════════════════════
    private static void InsertLineBreaks(WText text, string value)
    {
        if (text == null || value == null) return;

        var parts = value.Split('\n');
        text.Text = parts[0] ?? string.Empty;
        text.Space = SpaceProcessingModeValues.Preserve;

        if (parts.Length <= 1) return;

        var parentRun = text.Parent as Run;
        if (parentRun == null) return;

        OpenXmlElement anchor = text;
        for (int i = 1; i < parts.Length; i++)
        {
            var br = new Break();
            parentRun.InsertAfter(br, anchor);
            anchor = br;

            if (!string.IsNullOrEmpty(parts[i]))
            {
                var newText = new WText(parts[i])
                {
                    Space = SpaceProcessingModeValues.Preserve
                };
                parentRun.InsertAfter(newText, anchor);
                anchor = newText;
            }
        }
    }
}