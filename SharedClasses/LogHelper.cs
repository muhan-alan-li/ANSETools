using Serilog;
using Serilog.Events;

namespace SharedClasses;

public class LogHelper
{
    /// <summary>
    /// Configures the logging options for the duration of the program
    /// </summary>
    /// <para>
    /// There are three levels of logging available for this program
    /// 1. NONE
    /// 2. BASIC
    /// 3. VERBOSE
    /// 
    /// The NuGet package that provides logging (serilog) has six levels of
    /// logging available, for this program, we will use the following levels
    /// 1. VERBOSE
    /// 2. INFORMATION
    /// 3. WARNING
    /// 4. ERRORS
    ///
    /// If NONE is selected, no logs are written to file, and logs of only
    /// ERROR written to the console
    ///
    /// If BASIC is selected (-l), logs of INFORMATION, WARNING, and ERROR levels
    /// are written to file, and logs of ERROR level are written to the console
    ///
    /// If VERBOSE is selected (-v), logs of VERBOSE, INFORMATION, WARNING, and ERROR
    /// levels are written to file, and logs of ERROR and WARNING level are written
    /// to the console
    /// </para>
    /// <param name="options">
    /// Command line options object that was parsed from the command line input
    /// </param>
    public static void SetLogging(CLIOptions options)
    {
        // Write log level to the console
        string logFile = $"{DateTime.Now.ToString("hhmmss-")}.log";
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.Console()
            .CreateLogger();
        
        if (options.Verbose)
        {
            Log.Information("Log level set to: VERBOSE");
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Verbose()
                .WriteTo.File(logFile, rollingInterval: RollingInterval.Hour)
                .WriteTo.Console(restrictedToMinimumLevel: LogEventLevel.Warning)
                .CreateLogger();
            Log.Information("---- VERBOSE INFO LOG ----");
        }
        else if (options.Logging)
        {
            Log.Information("Log level set to: BASIC");
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File(logFile, rollingInterval: RollingInterval.Hour)
                .WriteTo.Console(restrictedToMinimumLevel: LogEventLevel.Error)
                .CreateLogger();
            Log.Information("---- BASIC INFO LOG ----");
        }
        else
        {
            Log.Information("Log level set to: NONE");
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Error()
                .WriteTo.Console(restrictedToMinimumLevel: LogEventLevel.Error)
                .CreateLogger();
        }
    }
}