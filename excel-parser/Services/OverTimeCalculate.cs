using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Globalization;
using System.Text.RegularExpressions;

namespace excel_parser.Services
{
    public static class OverTimeCalculate
    {
        public static Dictionary<string, string> Excute(Dictionary<string, string> workGroupDic, string mainFilePath)
        {
            IWorkbook workbook = WorkbookFactory.Create(mainFilePath);
            ISheet sheet = workbook.GetSheet("sheet");
            Dictionary<string, string> finallDic = [];

            var persNumAndSumIndexADic = new Dictionary<string, int>();
            var singlePersonNumber = string.Empty;
            var startSumIndex = -1;
            var cellIndex = 18;

            // Fill persNumAndSumIndexADic
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    var personNumber = FindPersonNumber(row);
                    if (!string.IsNullOrEmpty(personNumber))
                    {
                        singlePersonNumber = personNumber;
                    }
                    int startingCellIndex = GetColumnIndexByHeader(row, "مجموع");
                    if (startingCellIndex != -1)
                    {
                        startSumIndex = i;
                    }
                    if (startSumIndex != -1 && !string.IsNullOrEmpty(singlePersonNumber))
                    {
                        persNumAndSumIndexADic.TryAdd(singlePersonNumber, startSumIndex);

                        singlePersonNumber = string.Empty;
                        startSumIndex = -1;
                    }
                }
            }

            // Fill calculateModel list
            var calculateModels = new List<calculateModel>();
            foreach (var kvp in persNumAndSumIndexADic)
            {
                var PersonNumber = kvp.Key;


                var rowIndex = kvp.Value;

                if (rowIndex >= 0 && rowIndex <= sheet.LastRowNum)
                {
                    IRow row = sheet.GetRow(rowIndex);

                    if (row != null && cellIndex >= 0 && cellIndex < row.LastCellNum)
                    {
                        calculateModel model = new calculateModel
                        {
                            PersonNumber = PersonNumber,
                            WorkGroup = workGroupDic.ContainsKey(PersonNumber) ? workGroupDic[PersonNumber] : string.Empty,
                            TotalFinancialFunction = GetCellValue(row, cellIndex - 1),
                            OvertimeAuthorizedNormal = GetCellValue(row, cellIndex - 2),
                            OvertimeUnauthorizedNormal = GetCellValue(row, cellIndex - 3),
                            OvertimeAuthorizedHolidays = GetCellValue(row, cellIndex - 4),
                            OvertimeUnauthorizedHolidays = GetCellValue(row, cellIndex - 5),
                            OvertimeOnMission = GetCellValue(row, cellIndex - 6),
                            OvertimeSpecial = GetCellValue(row, cellIndex - 7),
                            FractionOfWorkDelay = GetCellValue(row, cellIndex - 8),
                            FractionOfWorkRush = GetCellValue(row, cellIndex - 9),
                            FractionOfWorkUnauthorizedExit = GetCellValue(row, cellIndex - 10),
                            FractionOfWorkAbsence = GetCellValue(row, cellIndex - 11),
                            FractionOfWorkSpecial = GetCellValue(row, cellIndex - 12),
                            WorkingAtNight = GetCellValue(row, cellIndex - 13),
                            AuthorizedDelay = GetCellValue(row, cellIndex - 14),
                            AuthorizedRush = GetCellValue(row, cellIndex - 15),
                            AllMission = GetCellValue(row, cellIndex - 16),
                            UnpaidLeave = GetCellValue(row, cellIndex - 17),
                            PaidLeave = GetCellValue(row, cellIndex - 18),
                        };

                        calculateModels.Add(model);
                    }
                    else
                    {
                        Console.WriteLine($"Error accurd while calculate overtime(Invalid Cell Index: {cellIndex})");
                    }
                }
                else
                {
                    Console.WriteLine($"Error accurd while calculate overtime(Invalid Row Index: {rowIndex})");
                }

            }

            // Fill final dictionary that have personal number and overtime
            foreach (var model in calculateModels)
            {
                switch (model.WorkGroup)
                {
                    case "گرمدره":
                        var resDicGr = TehranProcessor(model);
                        if (resDicGr.Count > 0)
                        {
                            finallDic.Add(resDicGr.Keys.FirstOrDefault() ?? "", resDicGr.Values.FirstOrDefault() ?? "");
                        }
                        break;

                    case "تهران":
                        var resDic = TehranProcessor(model);
                        if (resDic.Count > 0)
                        {
                            finallDic.Add(resDic.Keys.FirstOrDefault() ?? "", resDic.Values.FirstOrDefault() ?? "");
                        }
                        break;
                    default:
                        break;
                }
            }

            return finallDic;
        }
        private static string FindPersonNumber(IRow row)
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
                            var res = match.Groups[1].Value;
                            return (res);
                        }
                    }

                }
            }

            return "";
        }

        private static int GetColumnIndexByHeader(IRow headerRow, string header)
        {
            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                if (headerRow.GetCell(i)?.ToString()?.Trim() == header)
                {
                    return i;
                }
            }
            return -1;
        }

        private static string GetCellValue(IRow row, int cellIndex)
        {
            ICell cell = row.GetCell(cellIndex);
            return cell?.ToString() ?? string.Empty;
        }

        private static Dictionary<string, string> TehranProcessor(calculateModel model)
        {
            var resDic = new Dictionary<string, string>();


            var mainOvertime = ConvertToTotalMinutes(model.OvertimeAuthorizedNormal) +
                               ConvertToTotalMinutes(model.OvertimeAuthorizedHolidays) +
                               ConvertToTotalMinutes(model.OvertimeOnMission);

            var delay = ConvertToTotalMinutes(model.FractionOfWorkDelay) - ConvertToTotalMinutes(model.OvertimeUnauthorizedNormal);

            if (delay >= 0)
            {
                var result = mainOvertime - delay - (ConvertToTotalMinutes(model.AuthorizedRush) * 3) -
                              (ConvertToTotalMinutes(model.FractionOfWorkUnauthorizedExit) * 3);
                resDic.Add(model.PersonNumber, result.ToString());
            }
            else
            {
                var rush = (ConvertToTotalMinutes(model.FractionOfWorkRush) * 3) - Math.Abs(delay);

                if (rush >= 0)
                {
                    var result = mainOvertime - rush - (ConvertToTotalMinutes(model.FractionOfWorkUnauthorizedExit) * 3);
                    resDic.Add(model.PersonNumber, result.ToString());
                }
                else
                {
                    var exit = (ConvertToTotalMinutes(model.FractionOfWorkUnauthorizedExit) * 3) - (Math.Abs(rush));
                    var result = mainOvertime - exit;

                    if (result > 0)
                    {
                        resDic.Add(model.PersonNumber, result.ToString());
                    }
                }
            }

            return resDic;
        }


        private static Dictionary<string, string> GarmDarehProcessor(calculateModel model)
        {
            var resDic = new Dictionary<string, string>();

            var mainOvertime = ConvertToTotalMinutes(model.OvertimeAuthorizedNormal) +
                               ConvertToTotalMinutes(model.OvertimeAuthorizedHolidays) +
                               ConvertToTotalMinutes(model.OvertimeOnMission);

            var delay = ConvertToTotalMinutes(model.FractionOfWorkDelay) - ConvertToTotalMinutes(model.OvertimeUnauthorizedNormal);

            // Calculate rushMain and exitMain using ConvertToTotalMinutes helper function
            var rushMain = Math.Max(ConvertToTotalMinutes(model.FractionOfWorkRush) - 5, 0);
            var exitMain = Math.Max(ConvertToTotalMinutes(model.FractionOfWorkUnauthorizedExit) - 3, 0);

            if (delay >= 0)
            {
                var result = mainOvertime - delay - (rushMain * 3) - (exitMain * 3);
                resDic.Add(model.PersonNumber, result.ToString());
            }
            else
            {
                var rush = (rushMain * 3) - Math.Abs(delay);

                if (rush >= 0)
                {
                    var result = mainOvertime - rush - (exitMain * 3);
                    resDic.Add(model.PersonNumber, result.ToString());
                }
                else
                {
                    var exit = (exitMain * 3) - Math.Abs(rush);
                    var result = mainOvertime - exit;

                    if (result > 0)
                    {
                        resDic.Add(model.PersonNumber, result.ToString());
                    }
                }
            }

            return resDic;
        }

        private static int ConvertToTotalMinutes(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return 0;
            }

            if (value.Contains(':'))
            {
                // Check if the value has the format 'hh:mm'
                if (TimeSpan.TryParseExact(value, @"h\:mm", CultureInfo.InvariantCulture, out var timeSpan))
                {
                    // Convert the time duration to total minutes
                    return (int)timeSpan.TotalMinutes;
                }
            }

            // If not in 'hh:mm' format or regular integer conversion, assume it's in hours and convert to minutes
            return int.TryParse(value, out var hours) ? hours * 60 : 0;
        }

    }

    public class calculateModel()
    {
        public string PersonNumber { get; set; } = string.Empty;
        public string WorkGroup { get; set; } = string.Empty;
        public string TotalFinancialFunction { get; set; } = string.Empty;
        public string OvertimeAuthorizedNormal { get; set; } = string.Empty;
        public string OvertimeUnauthorizedNormal { get; set; } = string.Empty;
        public string OvertimeAuthorizedHolidays { get; set; } = string.Empty;
        public string OvertimeUnauthorizedHolidays { get; set; } = string.Empty;
        public string OvertimeOnMission { get; set; } = string.Empty;
        public string OvertimeSpecial { get; set; } = string.Empty;
        public string FractionOfWorkDelay { get; set; } = string.Empty;
        public string FractionOfWorkRush { get; set; } = string.Empty;
        public string FractionOfWorkUnauthorizedExit { get; set; } = string.Empty;
        public string FractionOfWorkAbsence { get; set; } = string.Empty;
        public string FractionOfWorkSpecial { get; set; } = string.Empty;
        public string WorkingAtNight { get; set; } = string.Empty;
        public string AuthorizedDelay { get; set; } = string.Empty;
        public string AuthorizedRush { get; set; } = string.Empty;
        public string AllMission { get; set; } = string.Empty;
        public string UnpaidLeave { get; set; } = string.Empty;
        public string PaidLeave { get; set; } = string.Empty;
    }

}
