using excel_parser.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;

var host = Host.CreateDefaultBuilder().ConfigureServices((hostContext, services) =>
{
    Log.Logger = new LoggerConfiguration()
        .MinimumLevel.Debug()
        .WriteTo.Console()
        .CreateLogger();

    services.AddSingleton<ManagerService>();

}).UseSerilog()
.Build();

#region Getting info
Log.Information("HELLO DEAR HR");

////Log.Information(":لطفا مسیر، اکسل مربوط به کد پرسنلی و گروه کاری را وارد نمایید");
//Log.Information("please enter the path of personal excel:");
//string perNumAndwrGroupfilePath = Console.ReadLine() ?? "";
//if (!System.IO.File.Exists(perNumAndwrGroupfilePath))
//{
//    Log.Information("مسیر وارد شده اشتباه می باشد و یا فایل وجود ندار.");
//    return;
//}

////Log.Information("لطفا مسیر اکسل کارکرد را وارد نمایید:");
//Log.Information("please enter the path of main excel:");
//string mainFilePath = Console.ReadLine() ?? "";
//if (!System.IO.File.Exists(mainFilePath))
//{
//    Log.Information("مسیر وارد شده اشتباه می باشد و یا فایل وجود ندار.");
//    return;
//}

////Log.Information("لطفا مسیر، اکسل مربوط به شب کاری و جمعه کاری را وارد نمایید:");
//Log.Information("please enter the path of night work excel:");
//string holidaysWorkFilePath = Console.ReadLine() ?? "";
//if (!System.IO.File.Exists(holidaysWorkFilePath))
//{
//    Log.Information("مسیر وارد شده اشتباه می باشد و یا فایل وجود ندار.");
//    return;
//}

////Log.Information("لطفا مسیر، اکسل ویژه را وارد نمایید:");
//Log.Information("please enter the path of custom excel:");
//string specialFilePath = Console.ReadLine() ?? "";
//if (!System.IO.File.Exists(specialFilePath))
//{
//    Log.Information("مسیر وارد شده اشتباه می باشد و یا فایل وجود ندار.");
//    return;
//}
#endregion

using (var scope = host.Services.CreateScope())
{
    var services = scope.ServiceProvider;
    var dataService = services.GetRequiredService<ManagerService>();
    //dataService.Execute(perNumAndwrGroupfilePath, mainFilePath, holidaysWorkFilePath, specialFilePath);
    dataService.Execute("C:\\Users\\mhm\\Documents\\GitHub\\Excel-Parser\\excel-parser\\Files\\main\\perwork.xlsx", "C:\\Users\\mhm\\Documents\\GitHub\\Excel-Parser\\excel-parser\\Files\\main\\haji2.xlsx", "C:\\Users\\mhm\\Documents\\GitHub\\Excel-Parser\\excel-parser\\Files\\main\\nightwork.xlsx", "C:\\Users\\mhm\\Documents\\GitHub\\Excel-Parser\\excel-parser\\Files\\main\\custom.xls");
}

host.Run();
