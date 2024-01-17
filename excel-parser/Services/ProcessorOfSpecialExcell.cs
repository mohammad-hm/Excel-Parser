using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public static class ProcessorOfSpecialExcell
{
    public static IWorkbook Execute(string specialFilePath, IWorkbook workbook)
    {
        // Create a workbook object from the input file
        IWorkbook inputWorkbook = WorkbookFactory.Create(specialFilePath);
        // Get the first worksheet in the input workbook
        ISheet inputSheet = inputWorkbook.GetSheet("sheet");

        // Create a new worksheet in the output workbook
        ISheet outputSheet = workbook.GetSheet("Output");

        // Iterate over the rows in the input worksheet
        for (int i = 1; i <= inputSheet.LastRowNum; i++)
        {
            // Get the current row from the input worksheet
            IRow inputRow = inputSheet.GetRow(i);

            if (inputRow != null)
            {
                // Get the personal number from the input row
                string inputPersonNumber = inputRow.GetCell(0)?.ToString() ?? "";

                // Find the corresponding row in the output worksheet based on the personal number
                IRow outputRow = FindOutputRowByPersonNumber(outputSheet, inputPersonNumber);

                if (outputRow != null)
                {
                    // Get the values of the specific columns in the input row
                    // TotalFinancialFunction
                    string getCell2 = inputRow.GetCell(2)?.ToString() ?? "";
                    // DailyMission
                    string getCell6 = inputRow.GetCell(6)?.ToString() ?? "";
                    // FractionOfWorkAbcenc
                    string getCell3 = inputRow.GetCell(3)?.ToString() ?? "";
                    // Family
                    string getCell1 = inputRow.GetCell(1)?.ToString() ?? "";

                    // Create cells in the output row and write the values
                    // TotalFinancialFunction
                    outputRow.CreateCell(2).SetCellValue(getCell2);
                    // DailyMission
                    outputRow.CreateCell(7).SetCellValue(getCell6);
                    // FractionOfWorkAbcenc
                    outputRow.CreateCell(4).SetCellValue(getCell3);
                    // Family
                    outputRow.CreateCell(1).SetCellValue(getCell1);
                }
            }
        }

        return workbook;
    }

    private static IRow FindOutputRowByPersonNumber(ISheet outputSheet, string personNumber)
    {
        // Iterate over the rows in the output worksheet
        for (int i = 1; i <= outputSheet.LastRowNum; i++)
        {
            IRow outputRow = outputSheet.GetRow(i);

            if (outputRow != null)
            {
                // Get the personal number from the output row
                string outputPersonNumber = outputRow.GetCell(0)?.ToString() ?? "";

                // Check if the personal numbers match
                if (outputPersonNumber == personNumber)
                {
                    return outputRow;
                }
            }
        }

        // If no matching row is found, return null
        return null;
    }
}
