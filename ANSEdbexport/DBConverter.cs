using System.Drawing;
using Microsoft.Data.Sqlite;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Serilog;
using SharedClasses;

namespace ANSEdbexport
{
    public class DBConverter
    {
        /* ---- Class Properties ---- */
        private static Format _outFormat;       // output format (either .xlsx or .in)

        /// <summary>
        /// Constructor, sets the output format to xlsx by default
        /// </summary>
        public DBConverter()
        {
            _outFormat = Format.Xlsx;
        } 
            
        /// <summary>
        /// Enum with two options: xlsx and text, not used publically
        /// </summary>
        private enum Format {
            Xlsx,
            Text
        }
        
        /// <summary>
        /// Called from outside the function, begins running the conversion process
        /// </summary>
        /// <param name="options">The command line options/arguments in object form</param>
        public void RunConvert(CLIOptions options)
        {
            try
            {
                // Check arguments are valid and return the file paths that are needed
                string[] argStrs     = InputChecker.CheckArgs(options);
                string   configFile  = argStrs[0];
                string   inFileName  = argStrs[1];
                string   outFileName = argStrs[2];
            
                // parse the config file into an object
                var cfgObj = InputChecker.ParseConfig(configFile);

                // parse the output file and call the function according to output format
                switch (outFileName.Split('.').Last())
                {
                    case "xlsx":
                        _outFormat = Format.Xlsx;
                        break;
                    case "in":
                        _outFormat = Format.Text;
                        break;
                    default:
                        throw new ArgumentException("Error: Output file format is incorrect");
                }
            
                // open connection to db
                using (SqliteConnection connection = new SqliteConnection($"Data Source={inFileName};"))
                {
                    connection.Open();
                
                    // iterate through all tables in the db 
                    string getTablesQuery = "SELECT name FROM sqlite_master WHERE type='table';";
                    SqliteCommand getTablesCmd = new SqliteCommand(getTablesQuery, connection);
                    using (SqliteDataReader tableReader = getTablesCmd.ExecuteReader())
                    {
                        /*
                         * If output is xlsx, create new ExcelPackage to represent file,
                         * and run the function to process from db model to xlsx
                         */
                        if (_outFormat == Format.Xlsx)
                        {
                            ExcelPackage outTable = new ExcelPackage(new FileInfo(outFileName));
                            while (tableReader.Read())
                                ProcessTableToXlsx(tableReader["name"].ToString(), cfgObj, connection, outTable);
                        }
                        /*
                         * If output is text, append text to file to show the beginning of
                         * converted content and run the function to process from db model to text
                         */
                        else if (_outFormat == Format.Text)
                        {
                            File.AppendAllText(outFileName, $"----CONTENTS GENERATED FROM: {inFileName}----\n\n");
                            while (tableReader.Read())
                                ProcessTableToTxt(tableReader["name"].ToString(), cfgObj, connection, outFileName);
                            File.AppendAllText(outFileName, $"----END OF FILE----");
                        }
                        // If output format is not either of the above, something has gone wrong
                        else throw new ArgumentException("Error: Output file format is incorrect");
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Processes a single db table into text, this function will be called multiple times, once for each table
        /// </summary>
        /// <param name="dbTableName">Name of the db table to process</param>
        /// <param name="cfgObj">Object representing the config json</param>
        /// <param name="connection">Connection to the database</param>
        /// <param name="outFileName">Path to the output file</param>
        /// <exception cref="ArgumentException">Throws an exception if configuration file is null</exception>
        private static void ProcessTableToTxt(string dbTableName, dynamic cfgObj, SqliteConnection connection, string outFileName)
        {
            // get the "types" field from the config
            var typesCfg = cfgObj["types"];
            if (typesCfg == null) throw new ArgumentException("Error: Config file must contain type definitions");
            if (typesCfg[dbTableName] == null)
            {
                Log.Information($"Table {dbTableName, -30} | No configuration file found for this type, skip table");
                return;
            }

            /*
             * Check the config file and grab the special characters, these include:
             *
             * Comment character: default to '#' ---> # COMMENT
             * Delimiter (head):  default to '<'
             * Delimiter (tail):  default to '>' ---> <TYPE>
             * Splitter:          default to ';' ---> <TYPE>name;1;2;3.0;;;
             */ 
            string[] specChars   = InputChecker.CheckSpecialChars(cfgObj);
            string   commentChar = specChars[0];
            string   delimitHead = specChars[1];
            string   delimitTail = specChars[2];
            string   splitter    = specChars[3];

            File.AppendAllText(outFileName, $"\n{commentChar} Generating text from {dbTableName}...\n\n");
            
            string getDataQuery = $"SELECT * FROM {dbTableName}";
            SqliteCommand getDataCmd = new SqliteCommand(getDataQuery, connection);
            using (SqliteDataReader dataReader = getDataCmd.ExecuteReader())
            {
                // iterate through every row of the db table
                while (dataReader.Read())
                {
                    string outText = $"{delimitHead}{dbTableName.Trim()}{delimitTail}";

                    // iterate through all columns of one row in the db table, this gets converted to one line of text
                    for (int i = 0; i < dataReader.FieldCount; i++)
                    {
                        var colVal = dataReader.GetValue(i);
                        Type colType = dataReader.GetFieldType(i);
                        outText += $"{ParseCol(colVal, colType)}{splitter}";
                    }
                    File.AppendAllText(outFileName, $"{outText}\n");
                }
            }
        }

        /// <summary>
        /// Parses a column within a row (cell) and returns a string representation of its value
        /// </summary>
        /// <param name="val">The object value from db table</param>
        /// <param name="type">The type of the data ie: int, double, etc</param>
        /// <returns>A string representation of the value from the db</returns>
        private static string ParseCol(object? val, Type type)
        {
            string outStr = string.Empty;
            if (val == null) return outStr;
            
            if (type == typeof(int))
            {
                int valInt = 0;
                if (int.TryParse(val.ToString(), out valInt))
                    outStr = valInt.ToString();
            }
            else if (type == typeof(double))
            {
                double valDub = 0.0;
                if (double.TryParse(val.ToString(), out valDub))
                    outStr = valDub.ToString();
            }
            else if (type == typeof(bool))
            {
                outStr = (val.ToString() == "1") ? "TRUE" : "FALSE";
            }

            if (outStr == string.Empty) outStr = val.ToString().Trim();

            return outStr;
        }
        
        /// <summary>
        /// Processes a single db table in multiple rows of content in a .xlsx file, function will be called multiple times
        /// </summary>
        /// <param name="dbTableName">Name of the db table to process</param>
        /// <param name="cfgObj">Object representation of the configuration json</param>
        /// <param name="connection">Connection to the database</param>
        /// <param name="outTable">ExcelPackage representing the output .xlsx file</param>
        private static void ProcessTableToXlsx(string dbTableName, dynamic cfgObj, SqliteConnection connection, ExcelPackage outTable)
        {
            // get the "types" field from the config
            var typesCfg = cfgObj["types"];
            if (typesCfg == null) throw new ArgumentException("Error: Config file must contain type definitions");
            if (typesCfg[dbTableName] == null)
            {
                Log.Information($"Table {dbTableName, -30} | No configuration file found for this type, skip table");
                return;
            }
            
            // find the corresponding worksheet for this type
            string outSheetName = GetSheetName(cfgObj, dbTableName);
            Log.Information($"Table {dbTableName, -30} | Inputs written to worksheet {outSheetName}");
            
            // look for column offset in config file, default value = 1
            int colOffset = 1;
            if (cfgObj["options"]["col_offset"] != null) colOffset = (int) cfgObj["options"]["col_offset"];
            
            // look for sheet with sheetname, if it doesn't exist, create new sheet with sheetname
            if (outTable.Workbook.Worksheets[outSheetName] == null) SetupNewSheet(outSheetName, outTable, colOffset);
            ExcelWorksheet outSheet = outTable.Workbook.Worksheets[outSheetName];
            
            // write type header for readability
            WriteTypeHeader(dbTableName, outSheet, colOffset++, connection);
            
            // Iterate through all rows of db table
            string getDataQuery = $"SELECT * FROM {dbTableName}";
            SqliteCommand getDataCmd = new SqliteCommand(getDataQuery, connection);
            using (SqliteDataReader dataReader = getDataCmd.ExecuteReader())
            {
                // Get row number
                int currRow = outSheet.Dimension.Rows + 1;
                outSheet.InsertRow(currRow, 1);
                while (dataReader.Read())
                {
                    // Write typename to first cell
                    outSheet.Cells[currRow, colOffset].Value = dbTableName.Trim();
                    
                    // Write the fields to the rest of the row
                    WriteTypeDeclaration(dataReader, outSheet, colOffset, currRow);
                    
                    // Increment and go to next row
                    currRow++;
                }
            }
            outTable.Save();
        }

        /// <summary>
        /// Writes one row of the type header (highlighted in green) for legibility
        /// </summary>
        /// <param name="tableName">Name of the db table (and also the ANSE model type)</param>
        /// <param name="outSheet">Worksheet that data should be written to</param>
        /// <param name="colOffset">The column at which to begin writing data</param>
        /// <param name="connection">Connection to the database</param>
        private static void WriteTypeHeader(string tableName, ExcelWorksheet outSheet, int colOffset, SqliteConnection connection)
        {
            // insert two rows, set current row to the bottommost row
            int currRow = outSheet.Dimension.Rows + 2;
            outSheet.InsertRow(currRow - 1, 2);

            // label the current row with KEYWORD: {typename} for readability
            var typeNameCell = outSheet.Cells[currRow, (++colOffset)];
            typeNameCell.Value = $"KEYWORD: {tableName}";
            typeNameCell.Style.Font.Bold = true;
            ResizeColumn(colOffset, $"KEYWORD: {tableName}", outSheet);
            
            // get the corresponding type from db
            string getDataQuery = $"SELECT * FROM {tableName}";
            SqliteCommand getDataCmd = new SqliteCommand(getDataQuery, connection);
            using (SqliteDataReader dataReader = getDataCmd.ExecuteReader())
            {
                for (int i = 0; i < dataReader.FieldCount; i++)
                {
                    int currCol = (i + 1) + colOffset;
                    string fieldName = dataReader.GetName(i);
                    outSheet.Cells[currRow, currCol].Value = fieldName.Trim();

                    // resize column for readability
                    ResizeColumn(currCol, fieldName, outSheet);
                }
                // recolor row for readability
                outSheet.Row(currRow).Style.Fill.PatternType = ExcelFillStyle.Solid;
                outSheet.Row(currRow).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#CCFFCC"));
            }
        }

        /// <summary>
        /// Write one row of a single type declaration from the db to the .xlsx sheet
        /// </summary>
        /// <param name="dataReader">Data reader for reading from the database table</param>
        /// <param name="outSheet">Worksheet that the data should be written to</param>
        /// <param name="colOffset">Column offset from which data should be written</param>
        /// <param name="currRow">Current row fir data to be written to</param>
        private static void WriteTypeDeclaration(SqliteDataReader dataReader, ExcelWorksheet outSheet, int colOffset, int currRow)
        {
            Log.Verbose($"\t\tWriting data to Row {currRow, -4}");
            for (int i = 0; i < dataReader.FieldCount; i++)
            {
                if (dataReader.IsDBNull(i)) continue;
                
                // current column number
                int currCol = (i + 1) + colOffset;
                
                // get information about cell value in db
                // string colName = dataReader.GetName(i); CURRENTLY NOT IN USE
                var colVal = dataReader.GetValue(i);
                Type dataType = dataReader.GetFieldType(i);
                
                // get the reference to the cell that we are writing to
                ExcelRange cell = outSheet.Cells[currRow, currCol];
                
                // write value to cell in sheet
                WriteTypeField(cell, dataType, colVal);
                
                // resize cell for readability
                ResizeColumn(currCol, colVal.ToString(), outSheet);
            }
        }
        
        /// <summary>
        /// Takes one cell of value from the database and write it to one cell in the worksheet
        /// </summary>
        /// <param name="cell">Object that allows the function to interact with the cell</param>
        /// <param name="dataType">The datatype of the db value ie: int, double, etc</param>
        /// <param name="fieldVal">Value of the db field</param>
        private static void WriteTypeField(ExcelRange cell, Type dataType, object? fieldVal)
        {
            if (fieldVal == null)
            {
                cell.Value = string.Empty;
                return;
            }

            if (dataType == typeof(int))
            {
                int valInt = 0;
                if (int.TryParse(fieldVal.ToString(), out valInt))
                {
                    cell.Value = valInt;
                    return;
                }
            }
            else if (dataType == typeof(double))
            {
                double valDub = 0.0;
                if (double.TryParse(fieldVal.ToString(), out valDub))
                {
                    cell.Value = valDub;
                    return;
                }
            }
            else if (dataType == typeof(bool))
            {
                int valBool = 0;
                if (int.TryParse(fieldVal.ToString(), out valBool))
                {
                    cell.Value = (valBool == 0) ? "FALSE" : "TRUE";
                    return;
                }
            }
            // if data type is string or something unparsable
            cell.Value = fieldVal.ToString().Trim();

            var val = cell.Value;
            // log info
            Log.Verbose($"\t\t\tData from db table cast to {dataType} and written to cell {@cell}", val, cell);
        }

        /// <summary>
        /// Creates a new sheet and sets it up for further processing
        /// </summary>
        /// <param name="sheetName">Name of the worksheet to create</param>
        /// <param name="outTable">The excel table in which the new sheet should be created</param>
        /// <param name="colOffset">Column spacing before which data should be written</param>
        private static void SetupNewSheet(string sheetName, ExcelPackage outTable, int colOffset)
        {
            outTable.Workbook.Worksheets.Add(sheetName);
            var initCell = outTable.Workbook.Worksheets[sheetName].Cells[1, colOffset];
            initCell.Value = "COMMENT";
            
            // recolour comment cell for readability
            initCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            initCell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            
            // put the sheet name in row 1 
            outTable.Workbook.Worksheets[sheetName].Cells[1, 1 + colOffset].Value = sheetName;
            
            // write info to log
            Log.Verbose($"Created sheet: {sheetName}");
        }

        /// <summary>
        /// According to the type declaration, get the appropriate sheet for data to be written to
        /// </summary>
        /// <param name="cfgObj">Object representation of the configuration json</param>
        /// <param name="tableName">Name of the db table (also the name of the ANSE model type)</param>
        /// <returns></returns>
        private static string GetSheetName(dynamic cfgObj, string tableName)
        {
            string sheetName = "INPUTS - MISC";
            var groupings = cfgObj["grouping"];
            foreach (var category in groupings)
            {
                foreach (string type in groupings[category.Name])
                {
                    if (type == tableName)
                    {
                        sheetName = $"INPUTS - {category.Name}";
                        break;
                    }
                }
            }
            return sheetName;
        }

        /// <summary>
        /// Resize a column based on the text within the cell
        /// </summary>
        /// <param name="col">Column to resize</param>
        /// <param name="colText">Text within the column</param>
        /// <param name="ws">Worksheet where the column rests</param>
        private static void ResizeColumn(int col, string colText, ExcelWorksheet ws)
        {
            double currWidth = ws.Column(col).Width;
            double newWidth = string.IsNullOrEmpty(colText) ? 0 : colText.Trim().Length * 1.05;
            ws.Column(col).Width = double.Max(currWidth, newWidth);
        }
    }
}

