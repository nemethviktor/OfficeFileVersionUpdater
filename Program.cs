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
            "\u001b[2J\u001b[3J"); // cls -> https://www.reddit.com/r/csharp/comments/k8flpr/comment/gextslz/?utm_source=share&utm_medium=web2x&context=3
        Console.Clear();

        ParseArgs(args: args);
        string[] wordFiles = { }, excelFiles = { }, powerPointFiles = { };
        CollateFiles(folderToParse: _folderToParse);

        string officeVer = ProgramHelpers.GetOfficeVer(); // this is used for blacklist-clearing.

        if (wordFiles.Length > 0 && !_skipWord)
        {
            Console.WriteLine(value: "Starting Word files...");
            ProgramHelpers.ClearBlackList(whichApp: "Word", officeVer: officeVer);
            WordHandler wordUpdater = new();
            Application wordApp = wordUpdater.StartApp();

            foreach (string wordFile in wordFiles)
                wordUpdater.ProcessAndSaveFile(wordDoc: wordUpdater.OpenFile(fileNameWithPath: wordFile,
                    wordApp: wordApp));

            wordUpdater.QuitApp();
        }


        if (excelFiles.Length > 0 && !_skipExcel)
        {
            Console.WriteLine(value: "Starting Excel files...");
            ProgramHelpers.ClearBlackList(whichApp: "Excel", officeVer: officeVer);
            ExcelHandler excelUpdater = new();
            Microsoft.Office.Interop.Excel.Application excelApp = excelUpdater.StartApp();

            foreach (string excelFile in excelFiles)
                excelUpdater.ProcessAndSaveFile(excelWbk: excelUpdater.OpenFile(fileNameWithPath: excelFile,
                    excelApp: excelApp));

            excelUpdater.QuitApp();
        }


        if (powerPointFiles.Length > 0 && !_skipPowerPoint)
        {
            Console.WriteLine(value: "Starting PowerPoint files...");
            ProgramHelpers.ClearBlackList(whichApp: "PowerPoint", officeVer: officeVer);
            PowerPointHandler powerPointUpdater = new();
            Microsoft.Office.Interop.PowerPoint.Application powerPointApp = powerPointUpdater.StartApp();

            foreach (string powerPointFile in powerPointFiles)
                powerPointUpdater.ProcessAndSaveFile(powerPointPres: powerPointUpdater.OpenFile(
                    fileNameWithPath: powerPointFile,
                    powerPointApp: powerPointApp));

            powerPointUpdater.QuitApp();
        }


        ProgramHelpers.ExitWithMessage(exitReason: ExitReasons.Ok);


        void CollateFiles(string folderToParse)
        {
            if (!_skipWord)
            {
                Console.WriteLine(value: "Starting file collation in root of " + folderToParse);
                wordFiles = Directory.EnumerateFiles(path: folderToParse, searchPattern: "*.*",
                                          searchOption: SearchOption.AllDirectories)
                                     .Where(predicate: s => s.ToLower()
                                                             .EndsWith(value: ".doc") ||
                                                            s.ToLower()
                                                             .EndsWith(value: ".docx") ||
                                                            s.ToLower()
                                                             .EndsWith(value: ".docm"))
                                     .ToArray();
                Console.WriteLine(value: $"{wordFiles.Length} Word files.");
            }

            // Excel files don't seem to have versionings. They're either "old" or "new".
            // ...as such we don't need the x-files.
            if (!_skipExcel)
            {
                excelFiles = Directory.EnumerateFiles(path: folderToParse, searchPattern: "*.*",
                                           searchOption: SearchOption.AllDirectories)
                                      .Where(predicate: s => s.ToLower()
                                                              .EndsWith(value: ".xls"))
                                      .ToArray();
                Console.WriteLine(value: $"{excelFiles.Length} Excel files.");
            }

            // PowerPoint files don't seem to have versionings. They're either "old" or "new".
            // ...as such we don't need the x-files.
            if (!_skipPowerPoint)
            {
                powerPointFiles = Directory.EnumerateFiles(path: folderToParse, searchPattern: "*.*",
                                                searchOption: SearchOption.AllDirectories)
                                           .Where(predicate: s => s.ToLower()
                                                                   .EndsWith(value: ".ppt") ||
                                                                  s.ToLower()
                                                                   .EndsWith(value: ".pps"))
                                           .ToArray();
                Console.WriteLine(value: $"{powerPointFiles.Length} PowerPoint files.");
            }

            Console.WriteLine(value: "File collation done.");

            if (wordFiles.Length + excelFiles.Length + powerPointFiles.Length == 0)
            {
                Console.WriteLine(value: "Nothing to do.");
                ProgramHelpers.ExitWithMessage(exitReason: ExitReasons.Ok);
            }
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