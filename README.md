# Statement of Account Printing — Public Cemetery Lot Rental

A Windows Forms (.NET 8) desktop app that imports an Excel list of
cemetery lot rental accounts and prints **one Statement of Account per
row** — as filled **MS Word (.docx)** documents using a template, or as
a **PDF** with one page per row.

## Open in Visual Studio

1. Unzip this folder.
2. Double-click `StatementOfAccountPrinting.sln` (Visual Studio 2022).
3. Right-click the solution → **Restore NuGet Packages**.
4. Press **F5** to run.

## Workflow (4 buttons, left to right)

| Step | Button | What it does |
|------|--------|--------------|
| 1 | **Import Excel** | Reads `.xlsx` (Sheet 1, row 1 = headers) into a `List<CemeteryLotAccount>` shown in the grid. |
| 2 | **Create Word Template** | Generates a starter `WordTemplate.docx` with `{{Placeholders}}`. **Open it in MS Word and restyle freely** — fonts, logo, letterhead, tables — keep the `{{...}}` tokens. |
| 3 | **Generate Word (.docx)** | Picks the template, fills one copy per row, saves `SOA_<LotNumber>.docx` into a folder you choose. |
| 4 | **Print to PDF** | Generates one A4 page per row into a single PDF (page count = row count). |

## Expected Excel columns (Sheet 1)

| Col | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 |
|-----|---|---|---|---|---|---|---|---|---|------|
| | LotNumber | LesseeName | Address | ContactNumber | LotType | AreaSqm | LeaseStart | LeaseEnd | RatePerMonth | AmountDue |

## Easy customization points

- **Add a column** → 3 steps, all marked with `✏️` in code:
  1. Add a property in `CemeteryLotAccount.cs`.
  2. Map it in `MainForm.btnImportExcel_Click`.
  3. Add it to `WordTemplateService.ToPlaceholders()` (Word) and
     `MainForm.CreatePdf()` (PDF).
- **Change the template layout** → edit `WordTemplateService.BuildTemplateBody()`
  or restyle the generated .docx in Word.
- **Output file naming** → `StatementPrefix` const in `MainForm.cs`.

## Word placeholders

`{{LotNumber}} {{LesseeName}} {{Address}} {{ContactNumber}} {{LotType}}
 {{AreaSqm}} {{LeaseStart}} {{LeaseEnd}} {{RatePerMonth}} {{AmountDue}}
 {{DateGenerated}}`

No Microsoft Office installation is required at runtime (Open XML SDK).
Template styling is done by you in MS Word afterwards.

## NuGet packages

- ClosedXML 0.102.2 — read .xlsx without Excel installed
- DocumentFormat.OpenXml 3.1.0 — read/write .docx templates
- QuestPDF 2024.3.10 — PDF output (community licence)
