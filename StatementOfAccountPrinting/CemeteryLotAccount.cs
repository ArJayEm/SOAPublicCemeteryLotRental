namespace StatementOfAccountPrinting;

/// <summary>
/// One Excel row = one cemetery lot rental account.
/// ═══════════════════════════════════════════════════════════════════
///  HOW TO ADD / EDIT COLUMNS (3 easy steps):
///
///   1. Add a property below, e.g.:
///        public string NewField { get; set; } = string.Empty;
///
///   2. Map it in MainForm.btnImportExcel_Click:
///        NewField = row.Cell(11).GetString().Trim(),
///
///   3. Use it in WordTemplateService.ToPlaceholders() and in
///      MainForm.CreatePdf(), e.g.  [ "NewField" ] = a.NewField
///
///  Excel layout (Sheet 1, row 1 = headers — headers are ignored):
///   Col 1 LotNumber | 2 LesseeName | 3 Address | 4 ContactNumber |
///   5 LotType | 6 AreaSqm | 7 LeaseStart | 8 LeaseEnd |
///   9 RatePerMonth | 10 AmountDue
/// ═══════════════════════════════════════════════════════════════════
/// </summary>
public class CemeteryLotAccount
{
    //public string Date { get; set; } = string.Empty;
    public string LesseeName { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
    public string LotBlockNumber { get; set; } = string.Empty;
    public string AreaSqm { get; set; } = string.Empty;
    public string NameOfDeceased { get; set; } = string.Empty;
    public string RentalOrRenewalFee { get; set; } = string.Empty;
    public string Surcharge { get; set; } = string.Empty;
    public string Interest { get; set; } = string.Empty;
    public string TotalAmountDue { get; set; } = string.Empty;
    public string DateGenerated { get; set; } = string.Empty;
}
