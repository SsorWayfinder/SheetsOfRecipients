using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace SheetReader;

public class Sheet
{
    private WorkbookPart workbookPart;

    public Sheet(string path)
    {
        subawards = new Dictionary<String, int>();
        // DirectoryInfo target = System.IO.Directory.GetParent( System.IO.Directory.GetCurrentDirectory() + path);
        // Console.WriteLine(target.FullName);

        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, false))
        {
            workbookPart = spreadsheetDocument.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            List<Row> rows = sheetData.Elements<Row>().ToList();
            Row gOtherDirectCostsRow =
                (from row in rows
                 where getCellValue((Cell)row.FirstChild) == "G."
                 select row).First();

            // Seems likely i will never need to be greater than the index of the row that starts with "H."
            for (int i = Convert.ToInt32((uint)gOtherDirectCostsRow.RowIndex); i < rows.Count; i++)
            {
                var row = rows[i];

                if (row.Elements<Cell>().Count() > 1)
                {
                    List<Cell> cells = row.Elements<Cell>().ToList();
                    if (null != cells[1] && null != getCellValue(cells[1]))
                    {
                        var cell_1_value = getCellValue(cells[1]);
                        if (cell_1_value.StartsWith("Subaward:"))
                        {
                            String key;
                            var remainder = cell_1_value.Replace("Subaward:", "").Trim();
                            if(remainder.Count() > 0)
                            {
                                key = remainder;
                            }
                            else if (null != getCellValue(cells[2]))
                            {
                                key = getCellValue(cells[2]).Trim();
                            }
                            else
                            {
                                // This is a case that exists within the given files, but it is not clear how it should be handled
                                // Were this a real-world project I would ask for requirements for scenarios when the files to not fit the format
                                key = "UNKNOWN";

                            }
                            var value = Int32.Parse(getCellValue(cells[4]));
                            if (subawards.ContainsKey(key))
                            {
                                subawards[key] += value;
                            }
                            else
                            {
                                subawards.Add(key, value);
                            }

                        }

                    }

                }


            }

        }
    }

    public Dictionary<String, int> subawards { get; }

    private string? getCellValue(Cell cell)
    {
        string? value = cell.InnerText.Length > 0 ? cell.InnerText : null;
        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
        {
            // For shared strings, look up the value in the
            // shared strings table.
            var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            // If the shared string table is missing, something 
            // is wrong. Return the index that is in
            // the cell. Otherwise, look up the correct text in 
            // the table.
            if (stringTable != null)
            {
                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }
        }

        return value;
    }
}

