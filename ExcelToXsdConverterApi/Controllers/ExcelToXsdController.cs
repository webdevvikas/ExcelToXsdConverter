using ExcelToXsdConverterApi.Dal;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Xml.Linq;

namespace ExcelToXsdConverterApi.Controllers
{
    [ApiController]
    public class ExcelToXsdController : ControllerBase
    {
        private readonly DataHandler _dataHandler;
        public ExcelToXsdController(DataHandler dataHandler)
        {
            _dataHandler = dataHandler ?? throw new ArgumentNullException(nameof(dataHandler));
        }

        [HttpPost("GetDataFromExcel")]
        public async Task<IActionResult> GetDataFromExcel(IFormFile excelFile)
        {
            try
            {
                using (var stream = new MemoryStream())
                {
                    await excelFile.CopyToAsync(stream);
                    stream.Position = 0;
                    string xsdContent = await _dataHandler.ReadExcelFileAsync(stream);
                    return Content(xsdContent, "application/xml");
                }
            }
            catch (Exception ex) { return StatusCode(500, $"Error: {ex.Message}"); }
        }
    }
}
