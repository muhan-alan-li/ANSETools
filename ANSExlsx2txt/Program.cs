using OfficeOpenXml;
using Serilog;
using SharedClasses;

namespace ANSExlsx2txt;
class Program {
    /// <summary>
    /// Serves as the main entrance point to the program. 
    /// </summary>
    /// <para>
    /// Takes in the command line arguments, parses them, then calls the 
    /// ProcessInput function to begin the conversion from .xlsx to .in file.
    /// </para>
    /// <param name="args"> the arguments to the command line app </param>    
    public static void Main(string[] args)
    {
        /* 
         * Init non-commercial license for EPPlus
         * This library can be used by NRCan and any non-commercial organizations for free
         */
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        try
        {
            // Get arguments from CMD line
            CLIOptions options = InputChecker.ParseArgs(args);

            // Write help output to CMD line
            if (options.Help) throw new ArgumentException("Writing usage instructions...");
            
            // Set logging
            LogHelper.SetLogging(options);

            // Enter program
            XlsxConverter convertWorker = new XlsxConverter(options);
            convertWorker.RunConvert();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            if (ex is ArgumentException) Usage.ExcelExportUsage();
            return;
        }
        
        Console.WriteLine("Conversion completed!");
        Log.Information("----CONVERSION COMPLETED----");
        Log.CloseAndFlush();
    }
}