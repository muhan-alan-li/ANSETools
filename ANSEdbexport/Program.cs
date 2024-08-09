using OfficeOpenXml;
using Serilog;
using SharedClasses;

namespace ANSEdbexport;

class Program
{
    /// <summary>
    /// Main function of the program, calls the RunConvert function to export the db model
    /// </summary>
    /// <param name="args">Command line input arguments</param>
    /// <exception cref="ArgumentException">Throws an exception if any of the arguments are invalid</exception>
    public static void Main(string[] args)
    {
        try
        {
            // Get arguments from CMD line
            CLIOptions options = InputChecker.ParseArgs(args);
            
            // Set logging options
            LogHelper.SetLogging(options);

            // If help flag is selected, write help output to CMD line
            if (options.Help) throw new ArgumentException("Writing usage instructions...");
            
            // Set License for EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            // Enter program
            DBExporter exportWorker = new DBExporter(options);
            exportWorker.RunExport();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            if (ex is ArgumentException) Usage.DBExportUsage();
            return;
        }
        
        Console.WriteLine("Conversion completed!");
        Log.Information("----CONVERSION COMPLETED----");
        Log.CloseAndFlush();
    }
}