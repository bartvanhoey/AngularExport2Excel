using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Net.Mime;
using System.Threading.Tasks;

namespace AngularExport2Excel.Controllers
{
    [Route("api/[controller]")]
    public class ExportToExcelController : Controller
    {

        [HttpGet("[action]")]
        public IActionResult GetExcelFile()
        {
            var stream = new ExcelFileGenerator().Generate();
            return new ExcelFileResult(stream);
        }

        public class ExcelFileResult : IActionResult
        {
            public MemoryStream MemoryStream { get; set; }

            public ExcelFileResult(MemoryStream memoryStream) => MemoryStream = memoryStream;

            public Task ExecuteResultAsync(ActionContext context)
            {
                MemoryStream.Seek(0, SeekOrigin.Begin);
                var fileStreamResult = new FileStreamResult(MemoryStream, MediaTypeNames.Application.Octet);
                return fileStreamResult.ExecuteResultAsync(context);
            }
        }



        public class ExcelFileGenerator
        {
            public MemoryStream Generate()
            {
                var memoryStream = new MemoryStream();

                using (var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    var sheets = workbookPart.Workbook.AppendChild(new Sheets());

                    sheets.AppendChild(new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet 1"
                    });

                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                    var row1 = new Row();
                    row1.AppendChild(
                        new Cell
                        {
                            CellValue = new CellValue("Bart Van Hoey"),
                            DataType = CellValues.String
                        }
                        
                    );
                    sheetData.AppendChild(row1);

                    var row2 = new Row();
                    row2.AppendChild(
                        new Cell
                        {
                            CellValue = new CellValue("https://be.linkedin.com/in/bartvanhoey-dotnet-developer"),
                            DataType = CellValues.String
                        }
                    );
                    sheetData.AppendChild(row2);

                    var row3 = new Row();
                    row3.AppendChild(
                        new Cell
                        {
                            CellValue = new CellValue("https://github.com/bartvanhoey"),
                            DataType = CellValues.String
                        }
                    );
                    sheetData.AppendChild(row3);

                    return memoryStream;
                }
            }
        }

    }
}
