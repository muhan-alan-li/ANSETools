namespace SharedClasses
{
    /// <summary>
    /// An object that represents all the valid arguments to the command line app
    /// </summary>
    public class CLIOptions
    {
        public bool Append { get; set; }
        public bool Help { get; set; }
        public bool Logging { get; set; }
        public bool Verbose { get; set; }
        public string? ConfigFile { get; set; }
        public string? InputFile { get; set; }
        public string? OutputFile { get; set; }
        public string? Worksheet { get; set; }
    }
}