using Microsoft.Data.Sqlite;
using Serilog;
using SharedClasses;

namespace ANSEdbexport;

public class DBExporter
{
    private static dynamic cfgObj;
    private static SqliteConnection dbConn;
    private static string outPath;
    private static Format outFormat;

    private enum Format
    {
        Text,
        Xlsx
    }

    public DBExporter(CLIOptions options)
    {
        string cfgPath, dbPath;
        
        if (!string.IsNullOrEmpty(options.ConfigFile)) cfgPath = options.ConfigFile;
        else throw new ArgumentException("Error: Must include config file");

        if (!string.IsNullOrEmpty(options.InputFile)) dbPath = options.InputFile;
        else throw new ArgumentException("Error: Must include input db file");

        if (!string.IsNullOrEmpty(options.OutputFile)) outPath = options.OutputFile;
        else outPath = $"{options.InputFile.Split('.')[0]}.xlsx";
        
        cfgObj = InputChecker.ParseConfig(cfgPath);
        if (cfgObj == null) throw new ArgumentException("Error: Configuration file is null");

        dbConn = new SqliteConnection($"Data Source={dbPath};");
        if (dbConn == null) throw new ArgumentException("Error: Database connection is null");

        if (outPath.Split('.').Last() == "xlsx") outFormat = Format.Xlsx;
        else if (outPath.Split('.').Last() == "in") outFormat = Format.Text;
        else throw new ArgumentException("Error: Output file format must be .xlsx or .in");
    }

    public void RunExport()
    {
        try
        {
            if (outFormat == Format.Text)
            {
                ExportToText textExporter = new ExportToText(cfgObj, dbConn, outPath);
                textExporter.Export();
            }
            else if (outFormat == Format.Xlsx)
            {
                ExportToXlsx xlsxExporter = new ExportToXlsx(cfgObj, dbConn, outPath);
                xlsxExporter.Export();
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex.Message);
            throw;
        }
    }
}