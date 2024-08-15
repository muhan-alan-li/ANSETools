# ANSE Tools

## Description

This repository contains ANSE Input Generator and Database Export tools created for the Carbon Accounting Team at the Pacific Forestry Centre. The ANSE model is a software simulation model maintained by the Carbon Accounting Team for tracking and simulating the carbon stored within Harvested Wood Products. 

**XLSX to Text Conversion Tool**

The primary goal of the xlsx2txt tool is to take in a .xlsx excel worksheet and generate a .in text file. The .in text file can then be used to create an ANSE model within a sqlite database. 

**DB Export Tool**

The primary goal of the dbexport tool is to take in a .db model file and generate either a .xlsx worksheet or a .in text file. This can be considered the opposite of the tool above


## Download / Installation
    Currently, this repository does not contain the command line executables, which should be uploaded at a later date
To view the codebase itself, simply clone the repo, then open with your editor of choice. For C# oriented IDEs such as Visual Studio/Jetbrains Rider, it will be easiest to use the IDE's open project feature to select the `.sln` file located within the ANSE_Input_Gen folder.

Otherwise, simply download the executable for the corresponding operating system below, then refer to the [user manual](./USERGUIDE.md) for usage instructions.

| Operating System | XLSX to Text Tool | DB Export Tool |
| ---------------- | ----------------- | -------------- |
| MacOS Silicon    |                   |                |
| MacOS Intel      |                   |                |
| Win10 x64        |                   |                |
| Ubuntu x64       |                   |                |
| ***More to be added*** | | |

## Dependencies
The two software tools use 6 external packages that are available through the C# packet manager NuGet. The names, usage, and licenses for these software packages are listed below.

*To use the command line executables above, it is **NOT** necessary to install these packages.*

| Name | Version used in software (as of 06/2024) | License | Usage |
| - | - | --- | --- |
| [EPPlus](https://www.nuget.org/packages/EPPlus) | 7.13 | [Polyform Noncommercial License](https://www.nuget.org/packages/EPPlus/7.2.0/License) <br> **Required Notice:** [License Terms](https://polyformproject.org/licenses/noncommercial/1.0.0/) | This package is used for opening, and reading from any excel worksheets that are set as an input to the program | 
| [Microsoft.Data.Sqlite](https://www.nuget.org/packages/Microsoft.Data.Sqlite)| 8.0.8 | [MIT](https://licenses.nuget.org/MIT) | This package is used for connecting to the sqlite database and reading data from the model to be written to an Excel or Text file in the DB exporter tool|
| [Newtonsoft.Json](https://www.nuget.org/packages/Newtonsoft.Json) | 13.0.3 | [MIT](https://licenses.nuget.org/MIT) | This package is used for reading and parsing the `config.json` file which describes the ANSE model configuration for the program | 
| [Seilog](https://www.nuget.org/packages/Serilog) | 4.0.0 | [Apache 2.0](https://licenses.nuget.org/Apache-2.0) | This package is used to provide optional logging at multiple levels |
| [Seilog.Sinks.Console](https://www.nuget.org/packages/Serilog.Sinks.Console) | 6.0.0 | [Apache 2.0](https://licenses.nuget.org/Apache-2.0) | This package is used to allow for the log to write to the console |
| [Seilog.Sinks.File](https://www.nuget.org/packages/Serilog.Sinks.File) | 6.0.0 | [Apache 2.0](https://licenses.nuget.org/Apache-2.0) | This package is used to allow for the log to write to a specific file |

## User Guide

Although the command line tool itself contains a help flag that display useful information with regards to the application usage, it is recommended to refer to the [user manual](./USERGUIDE.md) for instructions on how to use the command line application.

If one is already familiar with the configuration and optional arguments to the command line program, the `.exe` file may be used like so.

1. Place the `.exe` file into the desired directory

2. Create a `config.json` file (instruction in user manual)

3. Run the tool and specify the configuration file and input file with the `-c` and `-i` flags respectively

4. The output will be a `.in` text file that can be found in the same directory as the `.exe` executable

To debug the application, it may be useful to activate basic/verbose logging, please refer to the [user manual](./USERGUIDE.md) for more information.