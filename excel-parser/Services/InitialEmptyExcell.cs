using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public static class InitialEmptyExcell
{
    public static IWorkbook Execute()
    {
        // Create a new workbook
        IWorkbook workbook = new XSSFWorkbook();

        // Create a new worksheet
        ISheet sheet = workbook.CreateSheet("Output");

        // Create a header row with the property names of the model class
        IRow headerRow = sheet.CreateRow(0);
        headerRow.CreateCell(0).SetCellValue("کد");
        headerRow.CreateCell(1).SetCellValue("نام و نام خانوادگی");
        headerRow.CreateCell(2).SetCellValue("کارکرد ورزانه");
        headerRow.CreateCell(3).SetCellValue("کارکرد کسر تعجیل");
        headerRow.CreateCell(4).SetCellValue("غیبت");
        headerRow.CreateCell(5).SetCellValue("کارکرد مرخصی بدون حقوق");
        headerRow.CreateCell(6).SetCellValue("کارکرد اضافه کاری");
        headerRow.CreateCell(7).SetCellValue("کارکرد ماموریت روزانه");
        headerRow.CreateCell(8).SetCellValue("کارکرد کسر خروج غیر مجاز");
        headerRow.CreateCell(9).SetCellValue("کارکرد کسر کار");
        headerRow.CreateCell(10).SetCellValue("کارکرد جمعه کاری");
        headerRow.CreateCell(11).SetCellValue("کارکرد شب کاری");
        headerRow.CreateCell(12).SetCellValue("کارکرد نوبت کاری 15%");

        // Optionally, you can adjust the column widths to fit the content
        for (int i = 0; i < 13; i++)
        {
            sheet.AutoSizeColumn(i);
        }

        return workbook;

    }
}
