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

Log.Information(":لطفا مسیر، اکسل مربوط به کد پرسنلی و گروه کاری را وارد نمایید");
string perNumAndwrGroupfilePath = Console.ReadLine() ?? "";
if (!System.IO.File.Exists(perNumAndwrGroupfilePath))
{
    Log.Information("مسیر وارد شده اشتباه می باشد و یا فایل وجود ندار.");
    return;
}

Log.Information("لطفا مسیر اکسل کارکرد را وارد نمایید:");
string mainFilePath = Console.ReadLine() ?? "";
if (!System.IO.File.Exists(mainFilePath))
{
    Log.Information("مسیر وارد شده اشتباه می باشد و یا فایل وجود ندار.");
    return;
}

Log.Information("لطفا مسیر، اکسل مربوط به شب کاری و جمعه کاری را وارد نمایید:");
string holidaysWorkFilePath = Console.ReadLine() ?? "";
if (!System.IO.File.Exists(holidaysWorkFilePath))
{
    Log.Information("مسیر وارد شده اشتباه می باشد و یا فایل وجود ندار.");
    return;
}

Log.Information("لطفا مسیر، اکسل ویژه را وارد نمایید:");
string specialFilePath = Console.ReadLine() ?? "";
if (!System.IO.File.Exists(specialFilePath))
{
    Log.Information("مسیر وارد شده اشتباه می باشد و یا فایل وجود ندار.");
    return;
}
#endregion

using (var scope = host.Services.CreateScope())
{
    var services = scope.ServiceProvider;
    var dataService = services.GetRequiredService<ManagerService>();
    dataService.Execute(perNumAndwrGroupfilePath, mainFilePath, holidaysWorkFilePath, specialFilePath);
}

host.Run();
