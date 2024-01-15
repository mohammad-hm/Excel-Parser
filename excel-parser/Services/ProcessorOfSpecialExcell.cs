using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public static class ProcessorOfSpecialExcell
{
    public static IWorkbook Execute(string specialFilePath, IWorkbook workbook)
    {
        // Create a workbook object from the input file
        IWorkbook inputWorkbook = WorkbookFactory.Create(specialFilePath);
        // Get the first worksheet in the input workbook
        ISheet inputSheet = inputWorkbook.GetSheet("Page1");

        // Create a new worksheet in the output workbook
        ISheet outputSheet = workbook.GetSheet("Output");

        var size = inputSheet.LastRowNum - 1;
        // Iterate over the rows in the input worksheet
        for (int i = 0; i <= size; i++)
        {
            // Get the current row
            IRow inputRow = inputSheet.GetRow(i+1);

            // Create a new row in the output worksheet
            IRow outputRow = outputSheet.GetRow(i+1);

            // Get the values of the specific columns in the input row
            string getCell2 = inputRow.GetCell(2).ToString() ?? "";
            string getCell4 = inputRow.GetCell(4).ToString() ?? "";
            string getCell7 = inputRow.GetCell(7).ToString() ?? "";

            // Create cells in the output row and write the values
            outputRow.CreateCell(2).SetCellValue(getCell2);
            outputRow.CreateCell(4).SetCellValue(getCell4);
            outputRow.CreateCell(7).SetCellValue(getCell7);
        }

        return workbook;
    }
}

