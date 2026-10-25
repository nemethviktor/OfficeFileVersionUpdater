using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using CommandLine;
using Microsoft.Office.Interop.Word;

namespace OfficeFileVersionUpdater;

internal class Program
{
    private static string _folderToParse;
    private static bool _skipWord, _skipExcel, _skipPowerPoint;

    private static void Main(string[] args)
    {
        Console.WriteLine(
            value:
            "[2J[3J"); // cls -> https://www.reddit.com/r/csharp/comments/k8flpr/comment/gextslz/?utm_source=share&utm_medium=web2x&context=3
        Console.Clear();

        ParseArgs(args: args);
        CollateFiles(folderToParse: _folderToParse, wordFiles: out List<string> wordFiles,
            excelFiles: out List<string> excelFiles, powerPointFiles: out List<string> powerPointFiles);

        string officeVer = ProgramHelpers.GetOfficeVer(); // this is used for blacklist-clearing.

        if (wordFiles.Count > 0)
        {
            Console.WriteLine(value: "Starting Word files...");
            ProgramHelpers.ClearBlackList(whichApp: "Word", officeVer: officeVer);
            WordHandler wordUpdater = new();
            Application wordApp = wordUpdater.StartApp();
            try
            {
                foreach (string wordFile in wordFiles)
                {
                    Document wordDoc = wordUpdater.OpenFile(fileNameWithPath: wordFile, wordApp: wordApp);
                    if (wordDoc != null) wordUpdater.ProcessAndSaveFile(wordDoc: wordDoc);
                }
            }
            finally
            {
                wordUpdater.QuitApp();
            }
        }

        if (excelFiles.Count > 0)
        {
            Console.WriteLine(value: "Starting Excel files...");
            ProgramHelpers.ClearBlackList(whichApp: "Excel", officeVer: officeVer);
            ExcelHandler excelUpdater = new();
            Microsoft.Office.Interop.Excel.Application excelApp = excelUpdater.StartApp();
            try
            {
                foreach (string excelFile in excelFiles)
                {
                    Microsoft.Office.Interop.Excel.Workbook excelWbk =
                        excelUpdater.OpenFile(fileNameWithPath: excelFile, excelApp: excelApp);
                    if (excelWbk != null) excelUpdater.ProcessAndSaveFile(excelWbk: excelWbk);
                }
            }
            finally
            {
                excelUpdater.QuitApp();
            }
        }

        if (powerPointFiles.Count > 0)
        {
            Console.WriteLine(value: "Starting PowerPoint files...");
            ProgramHelpers.ClearBlackList(whichApp: "PowerPoint", officeVer: officeVer);
            PowerPointHandler powerPointUpdater = new();
            Microsoft.Office.Interop.PowerPoint.Application powerPointApp = powerPointUpdater.StartApp();
            try
            {
                foreach (string powerPointFile in powerPointFiles)
                {
                    Microsoft.Office.Interop.PowerPoint.Presentation powerPointPres =
                        powerPointUpdater.OpenFile(fileNameWithPath: powerPointFile, powerPointApp: powerPointApp);
                    if (powerPointPres != null) powerPointUpdater.ProcessAndSaveFile(powerPointPres: powerPointPres);
                }
            }
            finally
            {
                powerPointUpdater.QuitApp();
            }
        }

        ProgramHelpers.ExitWithMessage(exitReason: ExitReasons.Ok);
    }


    /// <summary>
    ///     Walks the folder tree once and sorts the legacy-format files into per-app lists.
    /// </summary>
    private static void CollateFiles(string folderToParse,
        out List<string> wordFiles,
        out List<string> excelFiles,
        out List<string> powerPointFiles)
    {
        Console.WriteLine(value: "Starting file collation in root of " + folderToParse);
        wordFiles = new List<string>();
        excelFiles = new List<string>();
        powerPointFiles = new List<string>();

        // Excel & PowerPoint files don't seem to have versionings. They're either "old" or "new".
        // ...as such we don't need the x-files for those.
        foreach (string file in Directory.EnumerateFiles(path: folderToParse, searchPattern: "*.*",
                     searchOption: SearchOption.AllDirectories))
        {
            if (Path.GetFileName(path: file)
                    .StartsWith(value: "~$", comparisonType: StringComparison.Ordinal)) // Office lock files
                continue;

            switch (Path.GetExtension(path: file).ToLowerInvariant())
            {
                case ".doc":
                case ".docx":
                case ".docm":
                    if (!_skipWord) wordFiles.Add(item: file);
                    break;
                case ".xls":
                    if (!_skipExcel) excelFiles.Add(item: file);
                    break;
                case ".ppt":
                case ".pps":
                    if (!_skipPowerPoint) powerPointFiles.Add(item: file);
                    break;
            }
        }

        if (!_skipWord) Console.WriteLine(value: $"{wordFiles.Count} Word files.");
        if (!_skipExcel) Console.WriteLine(value: $"{excelFiles.Count} Excel files.");
        if (!_skipPowerPoint) Console.WriteLine(value: $"{powerPointFiles.Count} PowerPoint files.");
        Console.WriteLine(value: "File collation done.");

        if (wordFiles.Count + excelFiles.Count + powerPointFiles.Count == 0)
        {
            Console.WriteLine(value: "Nothing to do.");
            ProgramHelpers.ExitWithMessage(exitReason: ExitReasons.Ok);
        }
    }


    /// <summary>
    ///     This is also responsible for parsing the program arguments/parameters
    /// </summary>
    /// <param name="args"></param>
    private static void ParseArgs(string[] args)
    {
        Parser parser = new(configuration: settings => { settings.CaseSensitive = false; });
        ParserResult<Options> parseResult = parser.ParseArguments<Options>(args: args)
                                                  .WithParsed(action: o =>
                                                       {
                                                           _folderToParse =
                                                               o.FolderToParse.Replace(oldValue: "\"", newValue: "");

                                                           if (!Directory.Exists(path: _folderToParse))
                                                               ProgramHelpers.ExitWithMessage(
                                                                   exitReason: ExitReasons.InvalidFolder);

                                                           _skipWord = o.SkipWord;
                                                           _skipExcel = o.SkipExcel;
                                                           _skipPowerPoint = o.SkipPowerPoint;
                                                       }
                                                   );
        foreach (Error parseError in parseResult.Errors) Console.WriteLine(value: parseError);
        if (string.IsNullOrWhiteSpace(value: _folderToParse))
        {
            ProgramHelpers.ExitWithMessage(exitReason: parseResult.Errors.Any()
                ? ExitReasons.InvalidParametersSupplied
                : ExitReasons.NoFolderPassed); // bye bye
            // bye bye
        }
    }


    /// <summary>
    ///     List of Exit Reasons/Codes
    /// </summary>
    internal enum ExitReasons
    {
        Ok,
        NoFolderPassed,
        InvalidFolder,
        WordNotInstalled,
        ExcelNotInstalled,
        PowerpointNotInstalled,
        InvalidParametersSupplied
    }

    /// <summary>
    ///     This is responsible for parsing the program arguments/parameters
    /// </summary>
    [SuppressMessage(category: "ReSharper", checkId: "UnusedAutoPropertyAccessor.Local")]
    private class Options
    {
        [Option(shortName: 'f', longName: "folderToParse", Required = true,
            HelpText =
                """Folder to parse -- this is recursive so you only need to specify the top-level. Use double-quotes if the folder name contains spaces (e.g. OfficeFileVersionUpdater.exe -f "C:\someting something")""")]
        public string FolderToParse { get; set; }

        [Option(longName: "SkipWord", Required = false, HelpText = "Do NOT Parse Word files.")]
        public bool SkipWord { get; set; }

        [Option(longName: "SkipExcel", Required = false, HelpText = "Do NOT Parse Excel files.")]
        public bool SkipExcel { get; set; }

        [Option(longName: "SkipPowerPoint", Required = false, HelpText = "Do NOT Parse PowerPoint files.")]
        public bool SkipPowerPoint { get; set; }
    }
}