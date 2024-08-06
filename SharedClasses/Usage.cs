namespace SharedClasses;

public class Usage
{
    /// <summary>
    /// Writes the help message to the console
    /// </summary>
    public static void ExcelExportUsage()
    {
        Console.WriteLine(@"
            Script that formats ANSE input files from Excel spreadsheets.

            The keyword records and field values are first created manually on a worksheet
            in an Excel workbook. This script then uses the EPPlus package to open the
            workbook, scans the specified worksheet for rows containing a valid keyword,
            and writes the row contents out to a text file, adding the required delimiters
            in the process.    


            Usage : ANSExlsx2txt [-a] [-c <filename>] [-i <filename>] [-l] [-w <name>] [-o <filename>] [-v]
                    ANSExlsx2txt [-c <filename>] [-i <filename>] 
                    ANSExlsx2txt [-h]

            -a              or --append:            Append output to the designated output
                                                    file (instead of overwriting it)

            -c              or --config:            MANDATORY: Config file that helps parse the input
                                            
            -h              or --help:              View this help
    
            -i <filename>   or --input=<filename>:  MANDATORY: The input file to process

            -l              or --logging            Enable basic logging, which is otherwise 
                                                    disabled without this flag
    
            -o <filename>   or --output=<filename>: The location to put the output file

            -v              or --verbose            Enable verbose logging, if both this 
                                                    and regular logging flags are set, 
                                                    verbose logging takes precedence. If
                                                    neither is set, then logging is disabled
    
            -w <name>       or --worksheet=<name>:  Which worksheet in the input file to
                                                    draw data from. If unspecified, will 
                                                    default to using all worksheets that
                                                    contain the word ""input"" 
                                                    (not case sensitive) in their title

                                                    To specify multiple worksheets, please
                                                    enter one string, with each individual
                                                    sheet name separated by semicolons like
                                                    so: ""name1;name2;long name - 3, with many chars;name4""

            If the above parameters are not specified, the script will attempt to use
            hard-coded default values. If necessary, these values may be edited; they
            can be found at the top of the ProcessCommandLineArgs() function.

            Input files must be a .xlsx Excel spreadsheet, older .xls files can be converted
            using Excel by utilising the 'export' option.
        ");
    }
    
    /// <summary>
    /// Writes the help message to the console
    /// </summary>
    public static void DBExportUsage()
    {
        Console.WriteLine(@"
            Script that exports an ANSE database model to either a .in text file or a .xlsx Excel file.
 
            Keyword fields and values are read from the database, and each table is converted
            either to text or to xlsx. The EPPlus package is used to create a workbook and write
            to specified worksheets. The configuration file is used to determine which worksheets
            a database table of declarations will be written to.

            Usage : ANSEdb2xlsx [-c <filename>] [-i <filename>] [-l] [-o <filename>] [-v]
                    ANSEdb2xlsx [-c <filename>] [-i <filename>] 
                    ANSEdb2xlsx [-h]

            -c              or --config:            MANDATORY: Config file that helps parse the input
                                            
            -h              or --help:              View this help message
    
            -i <filename>   or --input=<filename>:  MANDATORY: The database model file to process

            -l              or --logging            Enable basic logging, which is otherwise 
                                                    disabled without this flag
    
            -o <filename>   or --output=<filename>: The location to put the output file.
                                                    
                                                    The output file can be either a
                                                    .xlsx file or a .in file. If not 
                                                    specified here, the program will 
                                                    default to creating a .xlsx file

            -v              or --verbose            Enable verbose logging, if both this 
                                                    and regular logging flags are set, 
                                                    verbose logging takes precedence. If
                                                    neither is set, then logging is disabled

            If the above parameters are not specified, the script will attempt to use
            hard-coded default values. If necessary, these values may be edited; they
            can be found at the top of the ProcessCommandLineArgs() function.
        ");
    }
}