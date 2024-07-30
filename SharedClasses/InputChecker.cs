using Newtonsoft.Json;

namespace SharedClasses
{
    public class InputChecker
    {
        /// <summary>
        /// Parses the config.json input and creates a dynamic object based on the config file
        /// </summary>
        /// <param name="configFile">String of the path to the config.json file</param>
        /// <returns>Dynamic object that allows the data inside the config file to be accessed</returns>
        /// <exception cref="ArgumentException">Will throw an argument if there is no config file</exception>
        public static dynamic ParseConfig(string configFile)
        {
            // Read JSON string from file
            string cfgStr = File.ReadAllText(configFile);
            var jsonSettings = new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Ignore
            };
            /*
             * A dynamic object is used since we cannot be sure of the
             * properties that is in the config object. The required fields
             * for the config object should be made available in documentation.
             *
             * Unfortunately this is clunky but I don't have a better solution right now.
             * If there were updates to ANSE or to this tool, the config file, the
             * logic in this app, as well as the documentation will have to be
             * updated.
             */
            var cfgObj = JsonConvert.DeserializeObject(cfgStr, jsonSettings);
            if (cfgObj != null) return cfgObj;

            throw new ArgumentException("Error: Config file is null");
        }
        
        /// <summary>
        /// Parses command line input and separates the arguments and any values assigned to arguments
        /// </summary>
        /// <param name="args">arguments from command line represented by an array of strings</param>
        /// <returns>A CommandLineOptions object that represents the arguments to the tool</returns>
        /// <exception cref="ArgumentException">Will throw an argument exception if an unknown argument is entered</exception>
        public static CLIOptions ParseArgs(string[] args)
        {
            CLIOptions options = new CLIOptions();

            for (int i = 0; i < args.Length; i++)
            {
                /*
                 * If argument contains "=", expect the full declaration of an arg
                 * eg: --worksheet="ALL", instead of -w ALL
                 */
                if (args[i].Contains('='))
                {
                    string curr = args[i].Trim().Split('=')[0];
                    string next = args[i].Trim().Split('=')[1];

                    switch(curr)
                    {
                        case "--config":
                            options.ConfigFile = next;
                            break;
                        case "--input":
                            options.InputFile = next;
                            break;
                        case "--output":
                            options.OutputFile = next;
                            break;
                        case "--worksheet":
                            options.Worksheet = next;
                            break;
                        default: throw new ArgumentException($"Unknown option {args[i]}");
                    }
                    continue;
                }

                // Otherwise, expect arguments with values to have values immediately follow the flag
                switch(args[i])
                {
                    case "-a":
                        options.Append = true;
                        break;
                    case "--append":
                        options.Append = true;
                        break;
                    case "-c":
                        options.ConfigFile = args[++i];
                        break;
                    case "-h":
                        options.Help = true;
                        break;
                    case "--help":
                        options.Append = true;
                        break;
                    case "-i":
                        options.InputFile = args[++i];
                        break;
                    case "-l":
                        options.Logging = true;
                        break;
                    case "--logging":
                        options.Append = true;
                        break;
                    case "-o":
                        options.OutputFile = args[++i];
                        break;
                    case "-v":
                        options.Verbose = true;
                        break;
                    case "--verbose":
                        options.Append = true;
                        break;
                    case "-w":
                        options.Worksheet = args[++i];
                        break;
                    default: throw new ArgumentException($"Unknown option {args[i]}");
                }
            }
            return options;
        }
    }
}

