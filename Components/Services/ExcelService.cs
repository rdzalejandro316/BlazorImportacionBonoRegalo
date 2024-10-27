using Syncfusion.XlsIO;

namespace importacionBono.Components.Services;

public class ExcelService
{
    public MemoryStream CreateExcel()
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Excel2010;

            var workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.IsGridLinesVisible = true;
            worksheet.Range["A1"].Text = "BONO_REGALO";
            worksheet.Range["B1"].Text = "VALOR";
            worksheet.Range["C1"].Text = "PORCENTAJE";
            worksheet.Range["A1:C1"].CellStyle.Font.Bold = true;

            using (MemoryStream stream = new MemoryStream())
            {
                workbook.SaveAs(stream);
                return stream;
            }
        }
    }

}