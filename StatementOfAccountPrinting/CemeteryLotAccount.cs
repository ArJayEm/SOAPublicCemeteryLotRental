namespace StatementOfAccountPrinting;

/// <summary>
/// One Excel row = one cemetery lot rental account.
/// ═══════════════════════════════════════════════════════════════════
///  HOW TO ADD / EDIT COLUMNS (3 easy steps):
///
///   1. Add a property below, e.g.:
///        public string NewField { get; set; } = string.Empty;
///
///   2. Map it in MainForm.toolStripButtonImportExcel_Click:
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
    public string LesseeName { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
    public string LotBlockNumber { get; set; } = string.Empty;
    public string AreaSqm { get; set; } = string.Empty;
    public string NameOfDeceased { get; set; } = string.Empty;
    public string LastPayment { get; set; } = string.Empty;
    public string DelinquentYear { get; set; } = string.Empty;
    public string RentalOrRenewalFee { get; set; } = string.Empty;
    public string Surcharge { get; set; } = string.Empty;
    public string Interest { get; set; } = string.Empty;
    public string TotalAmountDue { get; set; } = string.Empty;

    public static readonly Dictionary<string, string> ColumnHeaders = new()
    {
        [nameof(CemeteryLotAccount.LesseeName)] = "Lessee",
        [nameof(CemeteryLotAccount.NameOfDeceased)] = "Deceased",
        [nameof(CemeteryLotAccount.Address)] = "Address",
        [nameof(CemeteryLotAccount.LotBlockNumber)] = "Block / Lot",
        [nameof(CemeteryLotAccount.AreaSqm)] = "Area (sq.m.)",
        [nameof(CemeteryLotAccount.LastPayment)] = "Last Payment",
        [nameof(CemeteryLotAccount.DelinquentYear)] = "Delinquent Year",
        [nameof(CemeteryLotAccount.RentalOrRenewalFee)] = "Rental / Renewal Fee",
        [nameof(CemeteryLotAccount.Surcharge)] = "Surcharge",
        [nameof(CemeteryLotAccount.Interest)] = "Interest",
        [nameof(CemeteryLotAccount.TotalAmountDue)] = "Total Amount Due",
    };
}