using ANSExlsx2txt.Exceptions;
using SharedClasses;
using OfficeOpenXml;
using Serilog;
using System.Text.RegularExpressions;

namespace ANSExlsx2txt;

public class Converter
{
    /* ---- Private Members ---- */
    private static dynamic      cfgObj;
    private static string       outFile;
    private static ExcelPackage inTable;
    private static string       sheetsToProc;

    private static string       delimitHead;
    private static string       delimitTail;
    private static string       commentChar;
    private static string       splitterChar;
    private static decimal      minCrumb;

    private static string       commentMarker;
    private static string       keywordMarker;

    /// <summary>
    /// 
    /// </summary>
    /// <param name="options"></param>
    /// <exception cref="ArgumentException"></exception>
    /// <exception cref="ConfigurationException"></exception>
    public Converter(CLIOptions options)
    {
        /*
         * Start by setting up the following
         *      1. config object
         *      2. output file path
         *      3. excel package of the input excel table
         *      4. names of the worksheets to process
         */
        // Get the input .xlsx table
        if (!string.IsNullOrEmpty(options.InputFile))
        {
            if (options.InputFile.Split('.').Last() != "xlsx") 
                throw new ArgumentException("Error: Input file must have .xlsx extension");
            inTable = new ExcelPackage(new FileInfo(options.InputFile));
        }
        else throw new ArgumentException("Error: Must include input file");
        
        // Get the config file
        if (!string.IsNullOrEmpty(options.ConfigFile))
        {
            if (options.ConfigFile.Split('.').Last() != "json")
                throw new ArgumentException("Error: Config file must have .json extension");
            cfgObj = InputChecker.ParseConfig(options.ConfigFile);
        }
        else throw new ArgumentException("Error: Must include config file");
        
        // Get the output file
        if (!string.IsNullOrEmpty(options.OutputFile)) outFile = options.OutputFile;
        else outFile = $"{Path.GetFileNameWithoutExtension(options.InputFile)}.in";
        
        // Get the worksheets to process
        if (!string.IsNullOrEmpty(options.Worksheet)) sheetsToProc = options.Worksheet;
        else sheetsToProc = "ALL";

        // Clear output file if options.Append is false
        if (!options.Append) File.WriteAllText(outFile, string.Empty);
        
        // Log input args
        Log.Verbose("Input arguments in object form: {@InArgs}", options);
        
        /*
         * Next, get any special characters needed for text conversion
         */
        // Check config file has "options" section
        if (cfgObj["options"] == null) throw new ConfigurationException("Error: Configuration file must contain options field");
        
        // Check delimiter character is in options
        if (cfgObj["options"]["delimiter"] != null)
        {
            if (cfgObj["options"]["delimiter"] is string delim)
            {
                delimitHead = delim;
                delimitTail = delim;
            }
            else if (cfgObj["options"]["delimiter"][0] != null &&
                     cfgObj["options"]["delimiter"][1] != null)
            {
                delimitHead = cfgObj["options"]["delimiter"][0];
                delimitTail = cfgObj["options"]["delimiter"][1];
            }
        }
        else throw new ConfigurationException("Error: Configuration options must contain character for delimiters");
        
        // Check comment character is in options
        if (cfgObj["options"]["comment"].ToString() is string commChar) commentChar = commChar;
        else throw new ConfigurationException("Error: Configuration options must contain character for comments");
        
        // Check splitter character is in options
        if (cfgObj["options"]["splitter"].ToString() is string splitChar) splitterChar = splitChar;
        else throw new ConfigurationException("Error: Configuration options must contain character for splitter");

        if (cfgObj["options"]["min_crumb"] is decimal min) minCrumb = min;
        else throw new ConfigurationException("Error: Configuration options must contain decimal (or 0) for minimum crumb size");
        
        /*
         * Finally, get the special keywords to recognize within the sheet that marks either
         *      1. Comments
         *      2. Type declarations
         */
        if (cfgObj["worksheet"] == null) throw new ConfigurationException("Error: Configuration file must contain worksheet keywords field");

        if (cfgObj["worksheet"]["commentMarker"].ToString() is string commMark) commentMarker = commMark;
        else throw new ConfigurationException("Error: Configuration options must contain worksheet marker for comments");

        if (cfgObj["worksheet"]["keywordMarker"].ToString() is string keyMark) keywordMarker = keyMark;
        else throw new ConfigurationException("Error: Configuration options must contain worksheet marker for type declaration");
    }

    /// <summary>
    /// Get the values from the program arguments, then begin processing the input worksheet
    /// </summary>
    /// <exception cref="ArgumentException">Throws an exception if any unexpected arguments are found</exception>
    public void RunConvert()
    {
        // Assign input arguments to variables
        try
        {
            // Log metadata
            Log.Information("Running conversion tool...");
            Log.Verbose("Time: {Now}", DateTime.Now.ToString("yyyy-MM-dd"));

            ProcessTable();
        }
        catch (Exception ex)
        {
            Log.Warning("----CONVERSION EXITED----");
            Log.Error("{ErrMsg}", ex.Message);
            if (ex is ArgumentException)
                throw new ArgumentException(ex.Message);
        }
    }
    
    /// <summary>
    /// Processes an entire excel table (an .xlsx file)
    /// </summary>
    /// <para>
    /// This function will iterate through all the worksheets in an Excel file
    /// If no specific worksheet is named, the function will process all worksheets
    /// that contain the word INPUT (case-insensitive), otherwise, the function
    /// will process only the specified worksheets
    /// </para>
    /// <exception cref="ArgumentException"></exception>
    private static void ProcessTable()
    {
        if (sheetsToProc == "ALL")
        {
            Log.Information("Table imported from {inFile}", inTable.File.Name);
            Log.Information("PROCESSING ALL SHEETS");
            foreach (var worksheet in inTable.Workbook.Worksheets)
            {
                // Check if sheet is empty
                if (worksheet == null || worksheet.Dimension == null) continue;
                if (worksheet.Name.Trim().Substring(0,5).ToUpper() != "INPUT") continue;
                
                // Log sheet name
                Log.Information("----> Converting sheet: {wsName}", worksheet);
                ProcessWorksheet(worksheet);
            }
        }
        /*
         * If a specific worksheet string is given, select the
         * specific worksheets and iterate through them only
         */
        else
        {
            foreach(string sheetName in sheetsToProc.Split(';'))
            {
                // Skip empty name
                if (string.IsNullOrEmpty(sheetName)) continue;

                // Checking sheet exists
                ExcelWorksheet worksheet = inTable.Workbook.Worksheets[sheetName.Trim()];
                if (worksheet == null) throw new ArgumentException($"Error: invalid worksheet name {sheetName.Trim()}");

                // Log sheet name
                Log.Information("Table imported from {inName}", inTable.File.Name);
                Log.Information("PROCESSING SHEET: {wsName}", sheetName.Trim().ToUpper());
                Log.Information("----> Converting sheet: {wsName}", sheetName.Trim().ToUpper());
                ProcessWorksheet(worksheet);
            }
        }
    }
    
    /// <summary>
    /// Processes an individual worksheet, this function will iterate through each row of the worksheet and process the information on every row
    /// </summary>
    /// <para>
    /// NOTE: EPPlus ignores leading whitespaces in worksheets
    /// </para>
    /// <param name="ws">ExcelWorksheet object that represents the input worksheet</param>
    private static void ProcessWorksheet(ExcelWorksheet ws)
    {
        for (int r = 1; r <= ws.Dimension.Rows; r++)
        {
            // Skip whitespaces at the beginning of a row
            int c = 1;
            while (ws.Cells[r, c].Value == null && c < ws.Dimension.Columns) c++;
            // If row is only whitespace, skip to next row
            if (c >= ws.Dimension.Columns) continue;

            // Look at the first value that shows up
            if (ws.Cells[r, c].Value is string firstStr)
            {
                /*
                 * Check if row starts with "COMMENT"
                 * Any input within the row is considered a comment
                 */
                if (firstStr.Trim().ToUpper() == commentMarker)
                {
                    string commentText = $"{commentChar}";
                    for (int i = c + 1; i <= ws.Dimension.Columns; i++)
                    {
                        if (ws.Cells[r, i].Value is string cellText)
                        {
                            if (!string.IsNullOrEmpty(cellText)) commentText += $"\t{cellText}";
                        }
                    }
                    File.AppendAllText(outFile, $"{commentText.Trim()}\n");
                    Log.Information($"Row {r, -4} | Comment: {commentText.Trim()}");
                    continue;
                }

                /* 
                 * Check if row starts with "KEYWORD"
                 * If it does, we will skip the row, as it only contains 
                 * information meant to improve readability
                 */
                if (firstStr.Trim().ToUpper().Split(':')[0] == keywordMarker)
                {
                    string typeName = firstStr.Split(':')[1].Trim().ToUpper();
                    Log.Information($"Row {r, -4} | KEYWORD: {typeName}");
                    continue;
                }

                /*
                 * Row contains actual declaration, process into .in file
                 */
                Log.Information($"Row {r, -4} | Type: {firstStr.Trim().ToUpper()}");
                ProcessRow(r, c, firstStr.Trim().ToUpper(), ws);
            }
        }
    }

    /// <summary>
    /// Processes the information on a specific row in the inputted worksheet
    /// </summary>
    /// <param name="r">Row index</param>
    /// <param name="c">Column index</param>
    /// <param name="typeName">The type that is being declared on the row</param>
    /// <param name="ws">ExcelWorksheet class that represents the input worksheet</param>
    /// <exception cref="InvalidOperationException"></exception>
    private static void ProcessRow(int r, int c, string typeName, ExcelWorksheet ws)
    {
        // Type declaration is delimited by a specific character known as the delimiter
        string outText = $"{delimitHead}{typeName}{delimitTail}";

        /*
         * Check if type exists in the config file,
         * each type may have multiple declarations in
         * the cfg file
         *
         * Try applying each definition until one fits
         * perfectly with the definition in the input sheet
         */
        if (cfgObj["types"][typeName] != null)
        {
            var typeCfgArr = cfgObj["types"][typeName];
            foreach (var typeCfg in typeCfgArr)
            {
                string defStr = MatchRowToCfg(r, c + 1, ws, typeCfg);
                if (!string.IsNullOrEmpty(defStr))  // successfully converted
                {
                    outText += defStr;
                    Log.Information($"Row {r, -4} | \t\tOutput generated from worksheet: {outText}");
                    outText += "\n";
                    File.AppendAllText(outFile, outText);
                    return;
                }
            }
        }
        // Throw error if configuration file definition cannot be matched to the declaration in the worksheet
        Log.Error($"Row {r, -4} | ({ws.Name}) Configuration file does not contain definition for {typeName}");
        throw new InvalidOperationException($"Error: Configuration file does not contain definition " +
                                            $"for {typeName} declared on row {r}");
    }

    /// <summary>
    /// Matches the declaration in a row to the configuration provided for that declared object
    /// </summary>
    /// <param name="r">Row number to process</param>
    /// <param name="c">Column number of the first column</param>
    /// <param name="ws">Excel worksheet object</param>
    /// <param name="typeCfg">Configruation of the declared type on the row</param>
    /// <returns>
    /// An empty string if the declaration does not match the definition in the configuration
    /// The converted string if the declaration matches the definition in the configuration
    /// </returns>
    private static string MatchRowToCfg(int r, int c, ExcelWorksheet ws, dynamic typeCfg)
    {
        string outStr = "";
        int col = c;
        /*
         * Try parsing each cell, and append the parsed value
         * to the output text if parse is successful
         */
        try
        {
            foreach (var prop in typeCfg)
            {
                outStr += $"{MatchCellToCfg(r, col, ws, typeCfg[prop.Name])}{splitterChar}";
                if (col >= ws.Dimension.Columns) break;
                col++;
            }
            
            // Check subsequent columns in the row for values (these will end up being ignored)
            while (col < ws.Dimension.Columns)
            {
                if (ws.Cells[r, col].Value != null)
                    throw new Exception($"Error: Values in row {r} does not match definition in configuration, " +
                                        $"Cells beyond cell {GetColumnStr(col)}{r} should be empty");
                col++;
            }
        }
        /*
         * If a cell parsing exception is thrown, catch the exception,
         * write warnings to the log, then check if the type definition
         * is overloaded
         *
         * If the definition is overloaded, it is possible that the next
         * definition in the configuration file can be applied without causing
         * a cell parsing exception
         */
        catch (CellParsingException cpe)
        {
            Log.Warning($"Row {r, -4} | {cpe.Message}");
            Log.Warning($"Row {r, -4} | \t\tCell {GetColumnStr(col)}{r} cannot be parsed with the " +
                        $"current definition");
            Log.Warning($"Row {r, -4} | Input does not match type definition " +
                        $"in config file, checking for next overloaded definition on this type");
            outStr = "";
        }

        return outStr;
    }

    /// <summary>
    /// Parses the value within the cell, and returns a string that can be written to the .in file
    /// </summary>
    /// <para>
    /// If a cell is empty, the function will check the configuration to check
    /// whether the field has a default value. If there is a default value, this
    /// default value is applied instead. If the field does not have a default
    /// value, the function will check if the value is mandatory, if the value
    /// is not mandatory, the field is left empty. If the value is mandatory,
    /// an error is thrown.
    ///
    /// If the cell is not empty, the function will check the configuration
    /// for the intended type of the field, then attempt to parse the cell
    /// value to the type for the field. If the cell value cannot be parsed
    /// eg: 2.384 -> to be cast to an integer
    /// an error will be thrown.
    /// </para>
    /// <param name="r">Row index</param>
    /// <param name="c">Column index</param>
    /// <param name="ws">ExcelWorksheet class that represents the input worksheet</param>
    /// <param name="fieldProperties">An object that represents the properties of the type</param>
    /// <returns>A string that represents the value within the cell</returns>
    /// <exception cref="InvalidOperationException">An exception in thrown if the value within the cell is invalid</exception>
    private static string MatchCellToCfg(int r, int c, ExcelWorksheet ws, dynamic fieldProperties)
    {
        // get cell value
        var val = ws.Cells[r, c].Value;

        // If cell is empty, check if the property is mandatory/has default
        if (val == null)
        {
            bool hasDefault = false;
            if (fieldProperties["Default"] != null) hasDefault = !(string.IsNullOrEmpty(fieldProperties["Default"].ToString()));
            bool isMandatory = false;
            if (fieldProperties["Mandatory"] != null) isMandatory = (bool) fieldProperties["Mandatory"];

            if (hasDefault)
            {
                Log.Verbose($"Row {r,-4} | \t\tValue from cell {GetColumnStr(c)}{r} is empty, " +
                            $"use default value \"{fieldProperties["Default"].ToString().Trim()}\" instead");
                return fieldProperties["Default"].ToString().Trim();
            }
            if (!isMandatory)
            {
                Log.Verbose($"Row {r,-4} | \t\tValue from cell {GetColumnStr(c)}{r} is empty, " +
                            $"field is not mandatory, and thus left empty");
                return "";
            }
            throw new CellParsingException($"Error: null value in cell {GetColumnStr(c)}{r}");
        }
        
        // Cell has value, check if it conforms to type
        string parsedVal = "";
        switch (fieldProperties["Type"].ToString().Trim())
        {
            case "Boolean":
                string word = val.ToString().Trim().ToUpper();
                if (word == "FALSE" || word == "TRUE")
                {
                    parsedVal = (word == "TRUE") ? "1" : "0"; 
                    break;
                }
                throw new CellParsingException($"Error: non-boolean value in cell {GetColumnStr(c)}{r}");
            case "Double":
                // Ensure the value is not in scientific format
                double valNum = 0.0;

                if (double.TryParse(val.ToString().Trim(), out valNum))
                {
                    decimal valDec = RemoveCrumb((decimal)valNum, minCrumb);
                    if (Regex.IsMatch(val.ToString(), "[Ee]")) parsedVal = valDec.ToString();
                    else parsedVal = valDec.ToString();
                    break;
                }
                throw new CellParsingException($"Error: non-double value in cell {GetColumnStr(c)}{r}");
            case "Integer":
                int valInt = 0;
                if (int.TryParse(val.ToString().Trim(), out valInt))
                {
                    parsedVal = valInt.ToString();
                    break;
                }
                throw new CellParsingException($"Error: non-integer value in cell {GetColumnStr(c)}{r}");
            case "String":
                parsedVal = val.ToString().Trim();
                
                // throw exception to show user character is illegal
                if (parsedVal.Contains(commentChar))
                    throw new CellParsingException($"Error: Text must not contain special character {commentChar}");
                if (parsedVal.Contains(splitterChar))
                    throw new CellParsingException($"Error: Text must not contain special character {splitterChar}");
                break;
            case "Text":
                parsedVal = val.ToString().Trim();
                
                // throw exception to show user character is illegal
                if (parsedVal.Contains(commentChar))
                    throw new CellParsingException($"Error: Text must not contain special character {commentChar}");
                if (parsedVal.Contains(splitterChar))
                    throw new CellParsingException($"Error: Text must not contain special character {splitterChar}");
                break;
        }

        if (string.IsNullOrEmpty(parsedVal))
            throw new CellParsingException($"Error: value in cell ({r}, {c}) could not be parsed");
        
        Log.Verbose($"Row {r,-4} | \t\tValue from cell {GetColumnStr(c)}{r} " +
                    $"cast to {$"({fieldProperties["Type"]})", -9} {parsedVal}");

        return parsedVal;
    }

    /// <summary>
    /// Converts a column number int other corresponding string
    /// Eg: 1 -> A, 27 -> AA
    /// </summary>
    /// <param name="colNum">Integer representing the column number</param>
    /// <returns>Returns a string that corresponds to the column number</returns>
    private static string GetColumnStr(int colNum)
    {
        string columnName = "";

        while (colNum > 0)
        {
            int modulo = (colNum - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            colNum = (colNum - modulo) / 26;
        }

        return columnName;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="num"></param>
    /// <param name="crumb"></param>
    /// <returns></returns>
    private static decimal RemoveCrumb(decimal num, decimal crumb)
    {
        decimal roundedNum = Math.Round(num);
        return (Math.Abs(roundedNum - num) <= crumb) ? roundedNum : num;
    }
}