using System;
using System.Globalization;
using System.IO;
using System.Linq;
using CommandLine;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;

namespace OfficeFileVersionUpdater;

internal class Program
{
    private static string _folderToParse;

    private static void Main(string[] args)
    {
        Console.WriteLine(
            "\u001b[2J\u001b[3J"); // cls -> https://www.reddit.com/r/csharp/comments/k8flpr/comment/gextslz/?utm_source=share&utm_medium=web2x&context=3
        Console.Clear();

        parseArgs(args);
        collateFiles(_folderToParse);

        string officeVer = ProgramHelpers.GetOfficeVer(); // this is used for blacklist-clearing.
        string[] wordFiles, excelFiles, powerPointFiles;
        string newFileName = "";
        if (wordFiles.Length > 0)
        {
            Console.WriteLine("Starting Word files...");
            ProgramHelpers.clearBlackList("Word", officeVer);
            Microsoft.Office.Interop.Word.Application wordApp = getNewWordApp();

            foreach (string wordFile in wordFiles)
            {
                FileInfo fi = new(wordFile);
                if (!fi.Name.Contains("~$"))
                {
                    DateTime lastModified = getLastModifiedDT(wordFile);

                    Document wordDoc = null;
                    try
                    {
                        Console.WriteLine("Processing " + wordFile);
                        wordDoc = wordApp.Documents.Open(wordFile);
                    }
                    catch (Exception ex)
                        // System.Runtime.InteropServices.COMException -- generally this will be stuff like pre-Office97 files. 
                        // HRESULT: 0x800A03EC -- you most likely provided the wrong password.
                    {
                        Console.WriteLine("-- " + ex.Message);
                    }

                    if (wordDoc != null)
                    {
                        // check compatibility [aka extension-eqsue]
                        // sub-15 is basically "old" refer to https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdcompatibilitymode?view=word-pia
                        if (wordDoc.CompatibilityMode < 15)
                        {
                            if (fi.Extension.ToLower() == ".doc") // super old
                            {
                                bool savedOK = false;
                                try
                                {
                                    if (wordDoc.HasVBProject)
                                    {
                                        newFileName = wordFile + "m";
                                        wordDoc.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocumentMacroEnabled);
                                        // word saves in compatibility mode for some reason by default.
                                        wordDoc.Convert(); // update to current
                                        wordDoc.Save();
                                        savedOK = true;
                                    }
                                    else
                                    {
                                        newFileName = wordFile + "x";
                                        wordDoc.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument);
                                        // word saves in compatibility mode for some reason by default.
                                        wordDoc.Convert(); // update to current
                                        wordDoc.Save();
                                        savedOK = true;
                                    }

                                    wordDoc.Close();
                                    printSavedAs(newFileName,
                                        lastModified); // kinda fake here for lastModified but if we got this far then setting the TS should be ok.
                                }
                                catch
                                {
                                    Console.WriteLine("-- Save failed for " + wordFile);
                                }

                                if (savedOK)
                                    // kill old version file if applicable
                                    File.Delete(wordFile);
                            }
                            else // basically this is stuff like version 2007 and generally "early" x-format files.
                            {
                                newFileName = wordFile;
                                wordDoc.Convert(); // update to current
                                wordDoc.Save();
                                wordDoc.Close();
                                Console.WriteLine("-- Re-saved file as version 2016 (current).");
                            }

                            File.SetLastWriteTime(newFileName, lastModified);
                        }
                        // non-legacy // compatibility-mode file
                        else
                        {
                            wordDoc.Close(false);
                            Console.WriteLine("- Ignored up-to-date version file: " + wordFile);
                        }
                    }
                }
            }

            wordApp.Quit(false);
            Console.WriteLine("Word files done");
        }

        if (excelFiles.Length > 0)
        {
            Console.WriteLine("Starting Excel files...");
            ProgramHelpers.clearBlackList("Excel", officeVer);
            Microsoft.Office.Interop.Excel.Application excelApp = getNewExcelApp();

            foreach (string excelFile in excelFiles)
            {
                FileInfo fi = new(excelFile);
                if (!fi.Name.Contains("~$"))
                {
                    DateTime lastModified = getLastModifiedDT(excelFile);

                    Workbook excelWbk = null;
                    try
                    {
                        Console.WriteLine("Processing " + excelFile);
                        excelWbk = excelApp.Workbooks.Open(excelFile);
                    }
                    catch (Exception ex)
                        // System.Runtime.InteropServices.COMException -- generally this will be stuff like pre-Office97 files. 
                        // HRESULT: 0x800A03EC -- you most likely provided the wrong password.
                    {
                        Console.WriteLine("-- " + ex.Message);
                    }

                    if (excelWbk != null)
                    {
                        // check compatibility [aka extension-eqsue]
                        if (excelWbk.Excel8CompatibilityMode)
                        {
                            bool savedOK = false;
                            try
                            {
                                if (excelWbk.HasVBProject)
                                {
                                    newFileName = excelFile + "m";
                                    excelWbk.SaveAs(newFileName, XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                                    savedOK = true;
                                }
                                else
                                {
                                    newFileName = excelFile + "x";
                                    excelWbk.SaveAs(newFileName, XlFileFormat.xlOpenXMLWorkbook);
                                    savedOK = true;
                                }

                                excelWbk.Close(false);
                            }
                            catch
                            {
                                Console.WriteLine("-- Save failed for " + excelFile);
                            }

                            if (savedOK)
                            {
                                // kill old version file if applicable
                                File.Delete(excelFile);
                                File.SetLastWriteTime(newFileName, lastModified);
                                printSavedAs(newFileName,
                                    lastModified); // kinda fake here for lastModified but if we got this far then setting the TS should be ok.
                            }
                        }
                        else // current-version file.
                        {
                            excelWbk.Close(false);
                            Console.WriteLine("- Ignored up-to-date version file: " + excelWbk);
                        }
                    }
                }
            }

            excelApp.Quit();
            Console.WriteLine("Excel files done");
        }

        if (powerPointFiles.Length > 0)
        {
            Console.WriteLine("Starting PowerPoint files...");
            ProgramHelpers.clearBlackList("PowerPoint", officeVer);
            Microsoft.Office.Interop.PowerPoint.Application powerPointApp = getNewPowerPointApp();

            foreach (string powerPointFile in powerPointFiles)
            {
                FileInfo fi = new(powerPointFile);
                if (!fi.Name.Contains("~$") &&
                    (fi.Extension.ToLower() ==
                     ".ppt" ||
                     fi.Extension.ToLower() ==
                     ".pps")) // that's correct actually. As per https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open the whole logic is that a file is either a PPT (old-ver) or not. 
                {
                    DateTime lastModified = getLastModifiedDT(powerPointFile);

                    Presentation powerPointPres = null;
                    try
                    {
                        Console.WriteLine("Processing " + powerPointFile);
                        powerPointPres = powerPointApp.Presentations.Open(powerPointFile, MsoTriState.msoTrue,
                            MsoTriState.msoFalse, MsoTriState.msoFalse);
                    }
                    catch (Exception ex)
                        // System.Runtime.InteropServices.COMException -- generally this will be stuff like pre-Office97 files. 
                        // HRESULT: 0x800A03EC -- you most likely provided the wrong password.
                    {
                        Console.WriteLine("-- " + ex.Message);
                    }

                    if (powerPointPres != null)
                    {
                        bool savedOK = false;
                        try
                        {
                            if (powerPointPres.HasVBProject)
                            {
                                newFileName = powerPointFile + "m";
                                powerPointPres.SaveAs(newFileName,
                                    fi.Extension.ToLower() ==
                                    ".ppt"
                                        ? PpSaveAsFileType.ppSaveAsOpenXMLPresentationMacroEnabled
                                        : PpSaveAsFileType.ppSaveAsOpenXMLShowMacroEnabled);

                                savedOK = true;
                            }
                            else
                            {
                                newFileName = powerPointFile + "x";

                                powerPointPres.SaveAs(newFileName,
                                    fi.Extension.ToLower() ==
                                    ".ppt"
                                        ? PpSaveAsFileType.ppSaveAsOpenXMLPresentation
                                        : PpSaveAsFileType.ppSaveAsOpenXMLShow);
                                savedOK = true;
                            }

                            powerPointPres.Close();
                        }
                        catch
                        {
                            Console.WriteLine("-- Save failed for " + powerPointFile);
                        }

                        if (savedOK)
                        {
                            File.Delete(powerPointFile);
                            File.SetLastWriteTime(newFileName, lastModified);
                            printSavedAs(newFileName,
                                lastModified); // kinda fake here for lastModified but if we got this far then setting the TS should be ok.
                        }
                    }
                }
            }

            powerPointApp.Quit();
            Console.WriteLine("PowerPoint files done");
        }

        exitWithMessage(exitReasons.OK);

        DateTime getLastModifiedDT(string fileToCheck)
        {
            return File.GetLastWriteTime(fileToCheck);
        }

        void collateFiles(string folderToParse)
        {
            Console.WriteLine("Starting file collation in root of " + folderToParse);
            wordFiles = Directory.EnumerateFiles(folderToParse, "*.*", SearchOption.AllDirectories)
                                 .Where(s => s.ToLower()
                                              .EndsWith(".doc") ||
                                             s.ToLower()
                                              .EndsWith(".docx") ||
                                             s.ToLower()
                                              .EndsWith(".docm"))
                                 .ToArray();
            Console.WriteLine($"{wordFiles.Length} Word files.");

            // Excel files don't seem to have versionings. They're either "old" or "new".
            // ...as such we don't need the x-files.
            excelFiles = Directory.EnumerateFiles(folderToParse, "*.*", SearchOption.AllDirectories)
                                  .Where(s => s.ToLower()
                                               .EndsWith(".xls"))
                                  .ToArray();
            Console.WriteLine($"{excelFiles.Length} Excel files.");

            // PowerPoint files don't seem to have versionings. They're either "old" or "new".
            // ...as such we don't need the x-files.
            powerPointFiles = Directory.EnumerateFiles(folderToParse, "*.*", SearchOption.AllDirectories)
                                       .Where(s => s.ToLower()
                                                    .EndsWith(".ppt") ||
                                                   s.ToLower()
                                                    .EndsWith(".pps"))
                                       .ToArray();
            Console.WriteLine($"{powerPointFiles.Length} PowerPoint files.");
            Console.WriteLine("File collation done.");

            if (wordFiles.Length + excelFiles.Length + powerPointFiles.Length == 0)
            {
                Console.WriteLine("Nothing to do. Exiting.");
                exitWithMessage(exitReasons.OK);
            }
        }

        void printSavedAs(string fileNameToSaveAs,
            DateTime lastModified)
        {
            Console.WriteLine("-- Saved as " +
                              fileNameToSaveAs +
                              " w/ TS " +
                              lastModified.ToString(CultureInfo.CurrentCulture));
        }
    }

    /// <summary>
    ///     Fetches a new Word process
    /// </summary>
    /// <returns>A Word process</returns>
    private static Microsoft.Office.Interop.Word.Application getNewWordApp()
    {
        Microsoft.Office.Interop.Word.Application wordApp = null;
        try
        {
            wordApp = new Microsoft.Office.Interop.Word.Application
            {
                Options =
                {
                    UpdateLinksAtOpen = false
                },
                WindowState = WdWindowState.wdWindowStateMinimize,
                DisplayAlerts = WdAlertLevel.wdAlertsNone,
                Visible = true // I thoroughly hate having to put this to Visible but it saves so much headache if user can actually interact.
            };
        }
        catch (Exception e)
        {
            exitWithMessage(exitReasons.WORD_NOT_INSTALLED);
        }

        return wordApp;
    }

    /// <summary>
    ///     Fetches a new Excel process
    /// </summary>
    /// <returns>An Excel process</returns>
    private static Microsoft.Office.Interop.Excel.Application getNewExcelApp()
    {
        Microsoft.Office.Interop.Excel.Application excelApp = null;
        try
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                WindowState = XlWindowState.xlMinimized,
                AskToUpdateLinks = false,
                Visible = true
            };
        }
        catch
        {
            exitWithMessage(exitReasons.EXCEL_NOT_INSTALLED);
        }

        return excelApp;
    }

    /// <summary>
    ///     Fetches a new PowerPoint process
    /// </summary>
    /// <returns>A PowerPoint process</returns>
    private static Microsoft.Office.Interop.PowerPoint.Application getNewPowerPointApp()
    {
        Microsoft.Office.Interop.PowerPoint.Application powerPointApp = null;
        try
        {
            powerPointApp = new Microsoft.Office.Interop.PowerPoint.Application
            {
                WindowState = PpWindowState.ppWindowMinimized,
                Visible = MsoTriState.msoTrue
            };
        }
        catch
        {
            exitWithMessage(exitReasons.POWERPOINT_NOT_INSTALLED);
        }

        return powerPointApp;
    }


    private static void parseArgs(string[] args)
    {
        Parser.Default.ParseArguments<Options>(args)
              .WithParsed(o =>
               {
                   _folderToParse = o.folderToParse.Replace("\"", "");

                   if (!Directory.Exists(_folderToParse)) exitWithMessage(exitReasons.INVALID_FOLDER);
               });
        if (string.IsNullOrWhiteSpace(_folderToParse)) exitWithMessage(exitReasons.NO_FOLDER_PASSED); // bye bye
    }

    /// <summary>
    ///     Exits the program with the given code. Easier to manage in one spot tbh.
    /// </summary>
    /// <param name="exitReason"></param>
    private static void exitWithMessage(exitReasons exitReason)
    {
        switch (exitReason)
        {
            case exitReasons.OK:
                Console.WriteLine("All seems to be well.");
                break;
            case exitReasons.NO_FOLDER_PASSED:
                Console.WriteLine("No folder provided.");
                break;
            case exitReasons.INVALID_FOLDER:
                Console.WriteLine("The path doesn't exist.");
                break;
            case exitReasons.WORD_NOT_INSTALLED:
                Console.WriteLine("Word installation not found.");
                break;
            case exitReasons.EXCEL_NOT_INSTALLED:
                Console.WriteLine("Excel installation not found.");
                break;
            case exitReasons.POWERPOINT_NOT_INSTALLED:
                Console.WriteLine("Powerpoint installation not found.");
                break;
        }

        Console.WriteLine("Exiting.");
        Environment.Exit((int)exitReason);
    }


    /// <summary>
    ///     List of Exit Reasons/Codes
    /// </summary>
    internal enum exitReasons
    {
        OK,
        NO_FOLDER_PASSED,
        INVALID_FOLDER,
        WORD_NOT_INSTALLED,
        EXCEL_NOT_INSTALLED,
        POWERPOINT_NOT_INSTALLED
    }

    /// <summary>
    ///     This is responsible for parsing the program arguments/parameters
    /// </summary>
    public class Options
    {
        [Option('f', "folderToParse", Required = true,
            HelpText =
                "Folder to parse -- this is recursive so you only need to specify the top-level.")]
        public string folderToParse { get; set; }
    }
}