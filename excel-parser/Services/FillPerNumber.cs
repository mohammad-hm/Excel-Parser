using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public static class FillPerNumber
{
    public static IWorkbook Execute(string perNumAndwrGroupfilePath, IWorkbook workbook)
    {


        // Create a workbook object from the input file
        IWorkbook inputWorkbook = WorkbookFactory.Create(perNumAndwrGroupfilePath);
        // Get the first worksheet in the input workbook
        ISheet inputSheet = inputWorkbook.GetSheet("sheet");


        // Create a new worksheet in the output workbook
        ISheet outputSheet = workbook.GetSheet("Output");

        var size = inputSheet.LastRowNum - 1;
        // Iterate over the rows in the input worksheet
        for (int i = 0; i <= size; i++)
        {
            // Get the current row
            IRow inputRow = inputSheet.GetRow(i + 1);

            if (inputRow.Count() > 0)
            {
                // Create a new row in the output worksheet
                IRow outputRow = outputSheet.CreateRow(i + 1);
                // Get the values of the second column in the input row
                string persionNumber = inputRow.GetCell(1).ToString() ?? "";

                // Create cells in the output row and write the values
                outputRow.CreateCell(0).SetCellValue(persionNumber);
            }

        }

        return workbook;

    }
}

