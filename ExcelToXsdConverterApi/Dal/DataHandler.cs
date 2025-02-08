using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Xml.Linq;
using System.IO;
using System.IO.Pipes;

namespace ExcelToXsdConverterApi.Dal
{
    public class DataHandler
    {
        public async Task<string> ReadExcelFileAsync(Stream fileStream)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(fileStream))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null) { throw new Exception("No worksheet found in the Excel file."); }

                    int colCount = worksheet.Dimension.Columns;
                    int rowCount = worksheet.Dimension.Rows;

                    var headers = new string[colCount];
                    for (int col = 1; col <= colCount; col++)
                    {
                        headers[col - 1] = worksheet.Cells[1, col].Text.Trim();
                    }

                    int idColumnIndex = Array.FindIndex(headers, h => h.Equals("ID", StringComparison.OrdinalIgnoreCase));
                    if (idColumnIndex == -1) { throw new Exception("ID column not found in the Excel file."); }

                    var specialRecords = new List<Dictionary<string, string>>();
                    var recordCounts = new Dictionary<string, int>();
                    var allRecords = new List<(Dictionary<string, string> Record, int RowIndex)>();

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var record = new Dictionary<string, string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            string cellValue = worksheet.Cells[row, col].Text.Trim();
                            record[headers[col - 1]] = cellValue;
                        }

                        string idValue = record[headers[idColumnIndex]];
                        if (!string.IsNullOrEmpty(idValue))
                        {
                            if (!recordCounts.ContainsKey(idValue))
                            {
                                recordCounts[idValue] = 0;
                            }
                            recordCounts[idValue]++;
                            allRecords.Add((record, row));
                        }
                    }

                    foreach (var (record, rowIndex) in allRecords)
                    {
                        string idValue = record[headers[idColumnIndex]];
                        var fillColor = worksheet.Cells[rowIndex, idColumnIndex + 1].Style.Fill.BackgroundColor;
                        bool isGreen = !string.IsNullOrEmpty(fillColor.Rgb) && fillColor.Rgb.StartsWith("C6EFCE", StringComparison.OrdinalIgnoreCase);

                        if (recordCounts[idValue] > 1 || isGreen)
                        {
                            specialRecords.Add(record);
                        }
                    }

                    return GenerateXsdFormat(specialRecords);
                }
            }
            catch (Exception ex)
            {
                return $"error: {ex.Message}";
            }
        }

        private string GenerateXsdFormat(List<Dictionary<string, string>> records)
        {
            XNamespace xs = "http://www.w3.org/2001/XMLSchema";

            var schema = new XElement(xs + "schema",
                new XAttribute(XNamespace.Xmlns + "xs", xs),
                new XElement(xs + "element",
                    new XAttribute("name", "records"),
                    new XElement(xs + "complexType",
                        new XElement(xs + "sequence",
                            records.Select(record =>
                                new XElement(xs + "record",
                                    record.Select(field =>
                                        new XElement(field.Key, field.Value)
                                    )
                                )
                            )
                        )
                    )
                )
            );

            return schema.ToString();
        }
    }
}
