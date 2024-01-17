using System.Collections.Generic;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public static class NightHolidayWorkService
{
    public static IWorkbook Execute(string holidaysWorkFilePath, IWorkbook workbook)
    {
        // Create a workbook object from the input file
        IWorkbook inputNightWorkbook = WorkbookFactory.Create(holidaysWorkFilePath);
        // Get the night work worksheet in the input workbook
        ISheet inputNightSheet = inputNightWorkbook.GetSheet("sheetn");
        // Get the holiday work worksheet in the input workbook
        ISheet inputHolidaySheet = inputNightWorkbook.GetSheet("sheeth");

        // Create a new worksheet in the output workbook
        ISheet outputSheet = workbook.GetSheet("Output");

        // Fill night work on output workbook
        var nightWorkDic = NightWorkDic(inputNightSheet);
        ProcessWorkDictionary(nightWorkDic, outputSheet, 11);

        // Fill holiday work on output workbook
        var holidayWorkDic = HolidayWorkDic(inputHolidaySheet);
        ProcessWorkDictionary(holidayWorkDic, outputSheet, 10);


        return workbook;
    }

    private static void ProcessWorkDictionary(Dictionary<string, string> workDic, ISheet outputSheet, int columnIndex)
    {
        if (workDic.Count > 0)
        {
            // Iterate through the rows in the "Output" sheet
            for (int i = 1; i <= outputSheet.LastRowNum; i++) // Start from 1 to skip header row
            {
                IRow outputRow = outputSheet.GetRow(i);
                if (outputRow != null)
                {
                    // Check if the first cell in the row contains the person number
                    ICell personNumCell = outputRow.GetCell(0);
                    if (personNumCell != null && workDic.TryGetValue(personNumCell.ToString() ?? "", out var workSum))
                    {
                        // If key is present, update the specified cell with the corresponding value from the dictionary
                        ICell cell = outputRow.CreateCell(columnIndex);
                        cell.SetCellValue(workSum);
                    }
                    else
                    {
                        // If key not present, set the specified cell value to "0"
                        ICell cell = outputRow.CreateCell(columnIndex);
                        cell.SetCellValue("0");
                    }
                }
            }
        }
        else
        {
            // If workDic is empty, set all the specified cells in "Output" sheet to "0"
            for (int i = 1; i <= outputSheet.LastRowNum; i++) // Start from 1 to skip header row
            {
                IRow outputRow = outputSheet.GetRow(i);
                if (outputRow != null)
                {
                    ICell cell = outputRow.CreateCell(columnIndex);
                    cell.SetCellValue("0");
                }
            }
        }
    }

    // Fill a dictionary that its key is personal number and its value is night work
    private static Dictionary<string, string> NightWorkDic(ISheet inputNightSheet)
    {
        var resDic = new Dictionary<string, string>();
        // Iterate over the rows in the input worksheet
        for (int i = 0; i <= inputNightSheet.LastRowNum; i++)
        {
            // Get the current row
            IRow row = inputNightSheet.GetRow(i);
            if (row != null)
            {
                for (int j = 0; j < row.LastCellNum; j++)
                {
                    ICell cellValue = row.GetCell(j);
                    if (cellValue != null)
                    {
                        if (Regex.IsMatch(cellValue.ToString() ?? "", @"شماره پرسنلی\s*:\s*(\d+)"))
                        {
                            // Extract the number using regular expression
                            Match match = Regex.Match(cellValue.ToString() ?? "", @"شماره پرسنلی\s*:\s*(\d+)");
                            if (match.Success)
                            {
                                var personNum = match.Groups[1].Value;
                                ICell nightWorkSum = row.GetCell(j + 6);
                                var nightWorkString = "";
                                if (nightWorkSum != null && nightWorkSum.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(nightWorkSum))
                                {
                                    DateTime dateValue = nightWorkSum.DateCellValue;
                                    nightWorkString = dateValue.ToString("h:mm:ss tt"); // Format according to your requirements
                                    
                                }
         
                                if (!string.IsNullOrEmpty(personNum) && !string.IsNullOrEmpty(nightWorkString))
                                {
                                    resDic.Add(personNum, nightWorkString);
                                }
                            }
                        }
                    }
                }

            }
        }
        return resDic;
    }

    private static Dictionary<string, string> HolidayWorkDic(ISheet inputHolidaySheet)
    {
        var resDic = new Dictionary<string, string>();
        // Iterate over the rows in the input worksheet
        for (int i = 0; i <= inputHolidaySheet.LastRowNum; i++)
        {
            // Get the current row
            IRow row = inputHolidaySheet.GetRow(i);
            if (row != null)
            {
                for (int j = 0; j < row.LastCellNum; j++)
                {
                    ICell cellValue = row.GetCell(j);
                    if (cellValue != null)
                    {
                        if (Regex.IsMatch(cellValue.ToString() ?? "", @"شماره پرسنلی\s*:\s*(\d+)"))
                        {
                            // Extract the number using regular expression
                            Match match = Regex.Match(cellValue.ToString() ?? "", @"شماره پرسنلی\s*:\s*(\d+)");
                            if (match.Success)
                            {
                                var personNum = match.Groups[1].Value;
                                ICell holidayWorkSum = row.GetCell(j + 6);
                                var holidayWorkString = "";
                                if (holidayWorkSum != null && holidayWorkSum.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(holidayWorkSum))
                                {
                                    DateTime dateValue = holidayWorkSum.DateCellValue;
                                    holidayWorkString = dateValue.ToString("h:mm:ss tt"); // Format according to your requirements

                                }
                                if (!string.IsNullOrEmpty(personNum) && !string.IsNullOrEmpty(holidayWorkString))
                                {
                                    resDic.Add(personNum, holidayWorkString);
                                }
                            }
                        }
                    }
                }
            }
        }
        return resDic;
    }
}
