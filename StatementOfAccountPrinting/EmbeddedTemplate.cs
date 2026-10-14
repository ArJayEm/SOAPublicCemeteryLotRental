using System;
using System.IO;
using System.Reflection;

public static class EmbeddedTemplate
{
    // Must match the file's logical name in the assembly:
    //   <RootNamespace>.<Folder>.<FileName>
    // e.g. StatementOfAccountPrinting.Templates.soa-public_cemetery_lot_rental.docx
    private const string ResourceName =
        "StatementOfAccountPrinting.Templates.soa-public_cemetery_lot_rental.docx";

    /// <summary>
    /// Copies the embedded .docx to a temp file and returns its path.
    /// The caller can open/modify it; delete it when done.
    /// </summary>
    public static string ExtractToTemp()
    {
        var asm = Assembly.GetExecutingAssembly();

        using var stream = asm.GetManifestResourceStream(ResourceName)
            ?? throw new FileNotFoundException(
                $"Embedded resource '{ResourceName}' not found. " +
                $"Available: {string.Join(", ", asm.GetManifestResourceNames())}");

        var tempPath = Path.Combine(
            Path.GetTempPath(),
            "soa_public_cemetery_lot_rental.docx");

        using var file = File.Create(tempPath);
        stream.CopyTo(file);
        return tempPath;
    }

    /// <summary>
    /// Copies the embedded .docx directly to a caller-chosen destination.
    /// </summary>
    public static void ExtractTo(string destinationPath)
    {
        var asm = Assembly.GetExecutingAssembly();

        using var stream = asm.GetManifestResourceStream(ResourceName)
            ?? throw new FileNotFoundException(
                $"Embedded resource '{ResourceName}' not found. " +
                $"Available: {string.Join(", ", asm.GetManifestResourceNames())}");

        // Ensure destination folder exists
        var dir = Path.GetDirectoryName(destinationPath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        using var file = File.Create(destinationPath);
        stream.CopyTo(file);
    }
}