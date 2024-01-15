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
        ISheet inputNightSheet = inputNightWorkbook.GetSheet("شبکاری");
        // Get the holiday work worksheet in the input workbook
        ISheet inputHolidaySheet = inputNightWorkbook.GetSheet("جمعه کاری");

        // Create a new worksheet in the output workbook
        ISheet outputSheet = workbook.GetSheet("Output");
        IRow outfputRow = outputSheet.GetRow(1);
        // Fill night work on output workbook
        var nightWorkDic = NightWorkDic(inputNightSheet);
        if (nightWorkDic.Count > 0)
        {
            // Iterate through the rows in the "Output" sheet
            for (int i = 0; i <= outputSheet.LastRowNum; i++)
            {
                IRow outputRow = outputSheet.GetRow(i);
                if (outputRow != null)
                {
                    // Check if the first cell in the row contains the person number
                    ICell personNumCell = outputRow.GetCell(0);
                    if (personNumCell != null && nightWorkDic.TryGetValue(personNumCell.ToString() ?? "", out var nightWorkSum))
                    {
                        ICell cell11 = outputRow.CreateCell(11);

                        cell11.SetCellValue(nightWorkSum);

                    }
                }
            }
        }

        // Fill holiday work on output workbook
        var holidayWorkDic = HolidayWorkDic(inputHolidaySheet);
        if (holidayWorkDic.Count > 0)
        {
            // Iterate through the rows in the "Output" sheet
            for (int i = 0; i <= outputSheet.LastRowNum; i++)
            {
                IRow outputRow = outputSheet.GetRow(i);
                if (outputRow != null)
                {
                    // Check if the first cell in the row contains the person number
                    ICell personNumCell = outputRow.GetCell(0);
                    if (personNumCell != null && holidayWorkDic.TryGetValue(personNumCell.ToString() ?? "", out var holidayWorkSum))
                    {
                        // If a match is found, update the fourth cell with the corresponding value from the dictionary
                        ICell cell10 = outputRow.CreateCell(10);

                        cell10.SetCellValue(holidayWorkSum);

                    }
                }
            }
        }

        return workbook;
    }

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
                                var nightWorkSum = row.GetCell(j + 5)?.ToString();
                                if (!string.IsNullOrEmpty(personNum) && !string.IsNullOrEmpty(nightWorkSum))
                                {
                                    resDic.Add(personNum, nightWorkSum);
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
                                var holidayWorkSum = row.GetCell(j + 5)?.ToString();
                                if (!string.IsNullOrEmpty(personNum) && !string.IsNullOrEmpty(holidayWorkSum))
                                {
                                    resDic.Add(personNum, holidayWorkSum);
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
