
using AddClientTestCodes.DTO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AddClientTestCodes.Helper
{
    public class ExcelHelper
    {

        public List<ExcelSheetDTO>? GetExcelFile(string customClientTestCode)
        {
            while (true)
            {
                Console.Write("Enter a file path (make sure to remove the quotation marks): ");
                var filePath = Console.ReadLine();

                if (filePath?.ToLower() == "exit")
                {
                    return null;
                }
                else if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
                {
                    return ReadExcel(@filePath, customClientTestCode);
                }
                else
                {
                    Console.WriteLine("Unable to find file.");
                    continue;
                }
            }
        }

        public List<ExcelSheetDTO> ReadExcel(string filePath, string customClientTestCode)
        {
            var list = new List<ExcelSheetDTO>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                Sheets sheetCollection = workbookPart.Workbook.GetFirstChild<Sheets>();

                foreach (Sheet sheet in sheetCollection.OfType<Sheet>())
                {
                    Worksheet worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                    for (int row = 1; row < sheetData.ChildElements.Count; row++)
                    {
                        ExcelSheetDTO record = new ExcelSheetDTO();

                        if (sheetData.ElementAt(row).ChildElements.Count == 3)
                        {
                            for (int cell = 0; cell < sheetData.ElementAt(row).ChildElements.Count; cell++)
                            {
                                Cell currentCell = (Cell)sheetData.ElementAt(row).ChildElements.ElementAt(cell);

                                if (currentCell.DataType != null)
                                {
                                    if (currentCell.DataType == CellValues.SharedString)
                                    {
                                        int id;

                                        if (int.TryParse(currentCell.InnerText, out id))
                                        {
                                            SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);

                                            if (item.InnerText != "" && item.InnerText != " ")
                                            {
                                                switch (cell)
                                                {
                                                    case 0:
                                                        record.PanelName = item.InnerText;
                                                        break;
                                                    case 2:
                                                        var genes = item.InnerText.Split("\n").ToList();
                                                        record.GenesInPanel = genes;
                                                        break;
                                                }

                                            }
                                            else
                                            {
                                                return list;
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    if (cell == 1 && currentCell.InnerText != "" && currentCell.InnerText != " ")
                                    {
                                        var numbers = Convert.ToInt32(currentCell.InnerText);
                                        record.GeneCount = numbers;

                                        var code = customClientTestCode;
                                        var custom = code + "-P" + numbers + "-";

                                        record.IncompleteTestCode = custom;
                                    }
                                    else
                                    {
                                        return list;
                                    }
                                }
                            }
                            list.Add(record);
                        }
                    }
                }
            }
            return list;
        }
    }
}
