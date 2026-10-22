using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using StatementOfAccountPrinting;
using WBreak = DocumentFormat.OpenXml.Wordprocessing.Break;
using WParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using WRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WRunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
// ── Aliases to prevent ambiguity with any other OpenXml namespace ──
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;
using WUnderline = DocumentFormat.OpenXml.Wordprocessing.Underline;
using WUnderlineValues = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues;

public static class WordTemplateService
{
    public const string TemplateFileName = "soa-public_cemetery_lot_rental.docx";
    public const string TemplateFolder = "Resources\\Templates";

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

    public static void GenerateCombinedWord(string outputPath, List<CemeteryLotAccount> accounts, string paperSize)
    {
        if (accounts == null || accounts.Count == 0)
            throw new ArgumentException("No accounts to generate.", nameof(accounts));

        var templatePath = ResolveTemplatePath();

        // 1) Copy the template to a writable temp file
        var tempPath = Path.Combine(Path.GetTempPath(), $"soa_{Guid.NewGuid():N}.docx");
        File.Copy(templatePath, tempPath, overwrite: true);

        try
        {
            // 2) Open the TEMP copy (never the original template)
            using (WordprocessingDocument templateDoc = WordprocessingDocument.Open(tempPath, isEditable: true))
            {
                var body = templateDoc.MainDocumentPart!.Document.Body!;

                foreach (var section in body.Elements<SectionProperties>())
                {
                    var pageSize = section.GetFirstChild<PageSize>();
                    if (pageSize == null)
                    {
                        pageSize = new PageSize();
                        section.Append(pageSize);
                    }

                    (pageSize.Width, pageSize.Height) = paperSize switch
                    {
                        "Legal" => (12240U, 20160U),
                        "A4" => (11906U, 16838U),
                        _ => (11906U, 16838U)
                    };
                }

                templateDoc.MainDocumentPart.Document.Save();
            }

            // 3) Now the temp copy has the right page size. Fill it.
            using var doc = WordprocessingDocument.Open(tempPath, isEditable: true);
            var main = doc.MainDocumentPart!;
            var body2 = main.Document.Body!;

            var originalChildren = body2.ChildElements
                .Select(e => e.CloneNode(true))
                .ToList();

            body2.RemoveAllChildren();

            for (int i = 0; i < accounts.Count; i++)
            {
                var account = accounts[i];
                var map = ToPlaceholders(account);

                foreach (var node in originalChildren)
                {
                    var clone = node.CloneNode(true);

                    var allParagraphs = clone.Descendants<WParagraph>().ToList();
                    if (clone is WParagraph self)
                        allParagraphs.Insert(0, self);

                    foreach (var para in allParagraphs)
                    {
                        var text = string.Concat(para.Descendants<WText>().Select(t => t.Text));
                        if (!text.Contains("{{")) continue;

                        if (para.Elements<WRun>().Count() > 1)
                            FlattenRuns(para);

                        ReplaceTokens(para, map);
                    }

                    body2.AppendChild(clone);
                }

                if (i < accounts.Count - 1)
                {
                    body2.AppendChild(new WParagraph(
                        new WRun(new WBreak { Type = BreakValues.Page })));
                }
            }

            main.Document.Save();

            // 4) Move the finished file to the requested output path
            if (File.Exists(outputPath)) File.Delete(outputPath);
            File.Move(tempPath, outputPath);
        }
        finally
        {
            // Clean up the temp file if anything failed
            if (File.Exists(tempPath))
            {
                try { File.Delete(tempPath); } catch { /* ignore */ }
            }
        }
    }

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
        foreach (var para in body.Descendants<WParagraph>())
        {
            FlattenRuns(para);
            ReplaceTokens(para, map);
        }

        // Headers
        foreach (var hp in main.HeaderParts)
        {
            foreach (var para in hp.Header.Descendants<WParagraph>())
            {
                FlattenRuns(para);
                ReplaceTokens(para, map);
            }
            hp.Header.Save();
        }

        // Footers
        foreach (var fp in main.FooterParts)
        {
            foreach (var para in fp.Footer.Descendants<WParagraph>())
            {
                FlattenRuns(para);
                ReplaceTokens(para, map);
            }
            fp.Footer.Save();
        }

        main.Document.Save();
    }

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

    public static void CreateTemplate(string outputPath, CemeteryLotAccount account)
    {
        if (account == null) throw new ArgumentNullException(nameof(account));

        CreateTemplate(outputPath);

        using var doc = WordprocessingDocument.Open(outputPath, isEditable: true);
        var main = doc.MainDocumentPart!;
        var body = main.Document.Body!;
        var map = ToPlaceholders(account);

        // Body
        foreach (var para in body.Descendants<WParagraph>())
        {
            FlattenRuns(para);
            ReplaceTokens(para, map);
        }

        // Headers
        foreach (var hp in main.HeaderParts)
        {
            foreach (var para in hp.Header.Descendants<WParagraph>())
            {
                FlattenRuns(para);
                ReplaceTokens(para, map);
            }
            hp.Header.Save();
        }

        // Footers
        foreach (var fp in main.FooterParts)
        {
            foreach (var para in fp.Footer.Descendants<WParagraph>())
            {
                FlattenRuns(para);
                ReplaceTokens(para, map);
            }
            fp.Footer.Save();
        }

        main.Document.Save();
    }

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
            ["LAST PAYMENT "] = (a.LastPayment ?? "").Trim(),
            ["DELINQUENT YEAR"] = (a.DelinquentYear ?? "").Trim(),
            ["TOTAL BASIC TAX-Final"] = a.RentalOrRenewalFee ?? "0.00",
            ["SURCHARGE"] = a.Surcharge ?? "0.00",
            ["INTEREST"] = a.Interest ?? "0.00",
            ["TOTAL DELIQUENCY"] = a.TotalAmountDue ?? "0.00",
        };
    }
    private static void FlattenRuns(WParagraph para)
    {
        if (para == null) return;

        var runs = para.Elements<WRun>().ToList();
        if (runs.Count == 0) return;

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
    private static readonly HashSet<string> UnderlineKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "DateGenerated",
        "NAME OF TAXPAYERS",
        "ADDRESS",
        "LotBlockNumber",
        "LOT AREA",
        "NAME OF DECEASED",
    };

    private static void ReplaceTokens(WParagraph para, Dictionary<string, string> map)
    {
        if (para == null || map == null) return;

        // 1) Full paragraph text
        var fullText = string.Concat(para.Descendants<WText>().Select(t => t.Text));
        if (string.IsNullOrEmpty(fullText) || !fullText.Contains("{{"))
            return;

        // 2) Replace every token in the full string
        var replaced = fullText;
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

        if (string.Equals(fullText, replaced, StringComparison.Ordinal))
            return;

        // 3) Snapshot the first run's properties so new runs inherit them
        var firstRun = para.Elements<WRun>().FirstOrDefault();
        var baseProps = firstRun?.RunProperties?.CloneNode(true) as WRunProperties;

        // 4) Clear all existing runs
        foreach (var r in para.Elements<WRun>().ToList())
            r.Remove();

        // 5) Rebuild
        if (underlineKey == null
            || !map.TryGetValue(underlineKey, out var value)
            || string.IsNullOrEmpty(value))
        {
            var run = new WRun();
            if (baseProps != null) run.AppendChild(baseProps.CloneNode(true));
            run.AppendChild(new WText(replaced) { Space = SpaceProcessingModeValues.Preserve });
            para.AppendChild(run);
            return;
        }

        int idx = replaced.IndexOf(value, StringComparison.Ordinal);
        if (idx < 0)
        {
            var run = new WRun();
            if (baseProps != null) run.AppendChild(baseProps.CloneNode(true));
            run.AppendChild(new WText(replaced) { Space = SpaceProcessingModeValues.Preserve });
            para.AppendChild(run);
            return;
        }

        var before = replaced.Substring(0, idx);
        var after = replaced.Substring(idx + value.Length);

        if (before.Length > 0)
        {
            var r = new WRun();
            if (baseProps != null) r.AppendChild(baseProps.CloneNode(true));
            r.AppendChild(new WText(before) { Space = SpaceProcessingModeValues.Preserve });
            para.AppendChild(r);
        }

        var valueRun = new WRun();
        var valueProps = baseProps?.CloneNode(true) as WRunProperties ?? new WRunProperties();
        foreach (var u in valueProps.Elements<WUnderline>().ToList()) u.Remove();
        valueProps.Append(new WUnderline { Val = WUnderlineValues.Single });

        valueRun.AppendChild(valueProps);
        valueRun.AppendChild(new WText(value) { Space = SpaceProcessingModeValues.Preserve });
        para.AppendChild(valueRun);

        if (after.Length > 0)
        {
            var r = new WRun();
            if (baseProps != null) r.AppendChild(baseProps.CloneNode(true));
            r.AppendChild(new WText(after) { Space = SpaceProcessingModeValues.Preserve });
            para.AppendChild(r);
        }
    }
    private static void UnderlineValueOnly(WText text, string value)
    {
        if (text == null || string.IsNullOrEmpty(value)) return;

        var run = text.Parent as WRun;
        if (run == null) return;

        var full = text.Text ?? "";
        int idx = full.IndexOf(value, StringComparison.Ordinal);
        if (idx < 0) return;

        var before = full.Substring(0, idx);
        var after = full.Substring(idx + value.Length);

        WRunProperties? originalProps = run.RunProperties;

        // before
        if (before.Length > 0)
        {
            var beforeRun = new WRun();
            if (originalProps != null)
                beforeRun.AppendChild(originalProps.CloneNode(true));
            beforeRun.AppendChild(new WText(before) { Space = SpaceProcessingModeValues.Preserve });
            run.InsertBeforeSelf(beforeRun);
        }

        // value (underlined)
        var valueRun = new WRun();
        var valueProps = originalProps != null
            ? (WRunProperties)originalProps.CloneNode(true)
            : new WRunProperties();

        foreach (var u in valueProps.Elements<WUnderline>().ToList())
            u.Remove();
        valueProps.Append(new WUnderline { Val = WUnderlineValues.Single });

        valueRun.AppendChild(valueProps);
        valueRun.AppendChild(new WText(value) { Space = SpaceProcessingModeValues.Preserve });
        run.InsertBeforeSelf(valueRun);

        // after
        if (after.Length > 0)
        {
            var afterRun = new WRun();
            if (originalProps != null)
                afterRun.AppendChild(originalProps.CloneNode(true));
            afterRun.AppendChild(new WText(after) { Space = SpaceProcessingModeValues.Preserve });
            run.InsertBeforeSelf(afterRun);
        }

        // remove original
        text.Remove();
        if (!run.Elements<WText>().Any()
            && !run.Elements<WBreak>().Any()
            && !run.Elements<Drawing>().Any())
        {
            run.Remove();
        }
    }
    private static void InsertLineBreaks(WText text, string value)
    {
        if (text == null || value == null) return;

        var parts = value.Split('\n');
        text.Text = parts[0] ?? string.Empty;
        text.Space = SpaceProcessingModeValues.Preserve;

        if (parts.Length <= 1) return;

        var parentRun = text.Parent as WRun;
        if (parentRun == null) return;

        OpenXmlElement anchor = text;
        for (int i = 1; i < parts.Length; i++)
        {
            var br = new WBreak();
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