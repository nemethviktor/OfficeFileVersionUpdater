using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace OfficeFileVersionUpdater;

internal class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("\u001b[2J\u001b[3J"); // cls -> https://www.reddit.com/r/csharp/comments/k8flpr/comment/gextslz/?utm_source=share&utm_medium=web2x&context=3
        Console.Clear();
        collateFiles(pathToScan: getPathToScan());
        string officeVer = ProgramHelpers.GetOfficeVer(); // this is used for blacklist-clearing.
        string[] wordFiles, excelFiles, powerPointFiles;
        string newFileName = "";
        if (wordFiles.Length > 0)
        {
            Console.WriteLine(value: "Starting Word files");
            ProgramHelpers.clearBlackList(whichApp: "Word", officeVer: officeVer);
            Application wordApp = new()
            {
                Options =
                {
                    UpdateLinksAtOpen = false
                },
                WindowState = WdWindowState.wdWindowStateMinimize,
                DisplayAlerts = WdAlertLevel.wdAlertsNone,
                Visible = true // I thoroughly hate having to put this to Visible but it saves so much headache if user can actually interact.
            };
            foreach (string wordFile in wordFiles)
            {
                FileInfo fi = new(fileName: wordFile);
                if (!fi.Name.Contains(value: "~$"))
                {
                    DateTime lastModified = getLastModifiedDT(fileToCheck: wordFile);

                    Document wordDoc = null;
                    try
                    {
                        Console.WriteLine(value: "Processing " + wordFile);
                        wordDoc = wordApp.Documents.Open(FileName: wordFile);
                    }
                    catch (Exception ex)
                        // System.Runtime.InteropServices.COMException -- generally this will be stuff like pre-Office97 files. 
                        // HRESULT: 0x800A03EC -- you most likely provided the wrong password.
                    {
                        Console.WriteLine(value: "-- " + ex.Message);
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
                                        wordDoc.SaveAs2(FileName: newFileName, FileFormat: WdSaveFormat.wdFormatXMLDocumentMacroEnabled);
                                        // word saves in compatibility mode for some reason by default.
                                        wordDoc.Convert(); // update to current
                                        wordDoc.Save();
                                        savedOK = true;
                                    }
                                    else
                                    {
                                        newFileName = wordFile + "x";
                                        wordDoc.SaveAs2(FileName: newFileName, FileFormat: WdSaveFormat.wdFormatXMLDocument);
                                        // word saves in compatibility mode for some reason by default.
                                        wordDoc.Convert(); // update to current
                                        wordDoc.Save();
                                        savedOK = true;
                                    }

                                    wordDoc.Close();
                                    printSavedAs(fileNameToSaveAs: newFileName,
                                                 lastModified: lastModified); // kinda fake here for lastModified but if we got this far then setting the TS should be ok.
                                }
                                catch
                                {
                                    Console.WriteLine(value: "-- Save failed for " + wordFile);
                                }

                                if (savedOK)
                                    // kill old version file if applicable
                                {
                                    File.Delete(path: wordFile);
                                }
                            }
                            else // basically this is stuff like version 2007 and generally "early" x-format files.
                            {
                                newFileName = wordFile;
                                wordDoc.Convert(); // update to current
                                wordDoc.Save();
                                wordDoc.Close();
                                Console.WriteLine(value: "-- Re-saved file as version 2016 (current).");
                            }

                            File.SetLastWriteTime(path: newFileName, lastWriteTime: lastModified);
                        }
                        // non-legacy // compatibility-mode file
                        else
                        {
                            wordDoc.Close(SaveChanges: false);
                            Console.WriteLine(value: "- Ignored up-to-date version file: " + wordFile);
                        }
                    }
                }
            }

            wordApp.Quit(SaveChanges: false);
            Console.WriteLine(value: "Word files done");
        }

        if (excelFiles.Length > 0)
        {
            Console.WriteLine(value: "Starting Excel files");
            ProgramHelpers.clearBlackList(whichApp: "Excel", officeVer: officeVer);
            Microsoft.Office.Interop.Excel.Application excelApp = new()
            {
                WindowState = XlWindowState.xlMinimized,
                AskToUpdateLinks = false,
                Visible = true
            };
            foreach (string excelFile in excelFiles)
            {
                FileInfo fi = new(fileName: excelFile);
                if (!fi.Name.Contains(value: "~$"))
                {
                    DateTime lastModified = getLastModifiedDT(fileToCheck: excelFile);

                    Workbook excelWbk = null;
                    try
                    {
                        Console.WriteLine(value: "Processing " + excelFile);
                        excelWbk = excelApp.Workbooks.Open(Filename: excelFile);
                    }
                    catch (Exception ex)
                        // System.Runtime.InteropServices.COMException -- generally this will be stuff like pre-Office97 files. 
                        // HRESULT: 0x800A03EC -- you most likely provided the wrong password.
                    {
                        Console.WriteLine(value: "-- " + ex.Message);
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
                                    excelWbk.SaveAs(Filename: newFileName, FileFormat: XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                                    savedOK = true;
                                }
                                else
                                {
                                    newFileName = excelFile + "x";
                                    excelWbk.SaveAs(Filename: newFileName, FileFormat: XlFileFormat.xlOpenXMLWorkbook);
                                    savedOK = true;
                                }

                                excelWbk.Close(SaveChanges: false);
                            }
                            catch
                            {
                                Console.WriteLine(value: "-- Save failed for " + excelFile);
                            }

                            if (savedOK)
                            {
                                // kill old version file if applicable
                                File.Delete(path: excelFile);
                                File.SetLastWriteTime(path: newFileName, lastWriteTime: lastModified);
                                printSavedAs(fileNameToSaveAs: newFileName,
                                             lastModified: lastModified); // kinda fake here for lastModified but if we got this far then setting the TS should be ok.
                            }
                        }
                        else // current-version file.
                        {
                            excelWbk.Close(SaveChanges: false);
                            Console.WriteLine(value: "- Ignored up-to-date version file: " + excelWbk);
                        }
                    }
                }
            }

            excelApp.Quit();
            Console.WriteLine(value: "Excel files done");
        }

        if (powerPointFiles.Length > 0)
        {
            Console.WriteLine(value: "Starting PowerPoint files");
            ProgramHelpers.clearBlackList(whichApp: "PowerPoint", officeVer: officeVer);
            Microsoft.Office.Interop.PowerPoint.Application powerPointApp = new()
            {
                WindowState = PpWindowState.ppWindowMinimized,
                Visible = MsoTriState.msoTrue
            };
            foreach (string powerPointFile in powerPointFiles)
            {
                FileInfo fi = new(fileName: powerPointFile);
                if (!fi.Name.Contains(value: "~$") &&
                    (fi.Extension.ToLower() ==
                     ".ppt" ||
                     fi.Extension.ToLower() ==
                     ".pps")) // that's correct actually. As per https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open the whole logic is that a file is either a PPT (old-ver) or not. 
                {
                    DateTime lastModified = getLastModifiedDT(fileToCheck: powerPointFile);

                    Presentation powerPointPres = null;
                    try
                    {
                        Console.WriteLine(value: "Processing " + powerPointFile);
                        powerPointPres = powerPointApp.Presentations.Open(FileName: powerPointFile, ReadOnly: MsoTriState.msoTrue,
                                                                          Untitled: MsoTriState.msoFalse, WithWindow: MsoTriState.msoFalse);
                    }
                    catch (Exception ex)
                        // System.Runtime.InteropServices.COMException -- generally this will be stuff like pre-Office97 files. 
                        // HRESULT: 0x800A03EC -- you most likely provided the wrong password.
                    {
                        Console.WriteLine(value: "-- " + ex.Message);
                    }

                    if (powerPointPres != null)
                    {
                        bool savedOK = false;
                        try
                        {
                            if (powerPointPres.HasVBProject)
                            {
                                newFileName = powerPointFile + "m";
                                powerPointPres.SaveAs(FileName: newFileName,
                                                      FileFormat: fi.Extension.ToLower() ==
                                                                  ".ppt"
                                                          ? PpSaveAsFileType.ppSaveAsOpenXMLPresentationMacroEnabled
                                                          : PpSaveAsFileType.ppSaveAsOpenXMLShowMacroEnabled);

                                savedOK = true;
                            }
                            else
                            {
                                newFileName = powerPointFile + "x";

                                powerPointPres.SaveAs(FileName: newFileName,
                                                      FileFormat: fi.Extension.ToLower() ==
                                                                  ".ppt"
                                                          ? PpSaveAsFileType.ppSaveAsOpenXMLPresentation
                                                          : PpSaveAsFileType.ppSaveAsOpenXMLShow);
                                savedOK = true;
                            }

                            powerPointPres.Close();
                        }
                        catch
                        {
                            Console.WriteLine(value: "-- Save failed for " + powerPointFile);
                        }

                        if (savedOK)
                        {
                            File.Delete(path: powerPointFile);
                            File.SetLastWriteTime(path: newFileName, lastWriteTime: lastModified);
                            printSavedAs(fileNameToSaveAs: newFileName,
                                         lastModified: lastModified); // kinda fake here for lastModified but if we got this far then setting the TS should be ok.
                        }
                    }
                }
            }

            powerPointApp.Quit();
            Console.WriteLine(value: "PowerPoint files done");
        }

        Environment.Exit(exitCode: 0);

        DateTime getLastModifiedDT(string fileToCheck)
        {
            return File.GetLastWriteTime(path: fileToCheck);
        }

        string getPathToScan()
        {
            string pathToScan = "C:\\temp\\";
            switch (args.Length)
            {
                case 0:
                    // nothing
                    break;
                case 1:
                    pathToScan =
                        args[0]
                           .Replace(oldValue: "\"",
                                    newValue: ""); // removes double quote - if the pathToScan has a space in it then args[0] comes through in odd ways.
                    break;
                default:
                    Console.WriteLine(
                        value: "Specify 0 or 1 count of folders to parse. Preferably 1, such as idk D:\\myfiles -- if you leave the arg blank it will default to " +
                               pathToScan);
                    Environment.Exit(exitCode: 1); // bye bye
                    break;
            }

            if (!Directory.Exists(path: pathToScan))
            {
                Console.WriteLine(value: pathToScan + " does not exist.");
                Environment.Exit(exitCode: 1); // bye bye
            }

            return pathToScan;
        }

        void collateFiles(string pathToScan)
        {
            Console.WriteLine(value: "Starting file collation in root of " + pathToScan);
            wordFiles = Directory.EnumerateFiles(path: pathToScan, searchPattern: "*.*", searchOption: SearchOption.AllDirectories)
                                 .Where(predicate: s => s.ToLower()
                                                         .EndsWith(value: ".doc") ||
                                                        s.ToLower()
                                                         .EndsWith(value: ".docx") ||
                                                        s.ToLower()
                                                         .EndsWith(value: ".docm"))
                                 .ToArray();
            Console.WriteLine(value: $"{wordFiles.Length} Word files.");

            // Excel files don't seem to have versionings. They're either "old" or "new".
            // ...as such we don't need the x-files.
            excelFiles = Directory.EnumerateFiles(path: pathToScan, searchPattern: "*.*", searchOption: SearchOption.AllDirectories)
                                  .Where(predicate: s => s.ToLower()
                                                          .EndsWith(value: ".xls"))
                                  .ToArray();
            Console.WriteLine(value: $"{excelFiles.Length} Excel files.");

            // PowerPoint files don't seem to have versionings. They're either "old" or "new".
            // ...as such we don't need the x-files.
            powerPointFiles = Directory.EnumerateFiles(path: pathToScan, searchPattern: "*.*", searchOption: SearchOption.AllDirectories)
                                       .Where(predicate: s => s.ToLower()
                                                               .EndsWith(value: ".ppt") ||
                                                              s.ToLower()
                                                               .EndsWith(value: ".pps"))
                                       .ToArray();
            Console.WriteLine(value: $"{powerPointFiles.Length} PowerPoint files.");
            Console.WriteLine(value: "File collation done.");

            if (wordFiles.Length + excelFiles.Length + powerPointFiles.Length == 0)
            {
                Console.WriteLine(value: "Nothing to do. Exiting.");
                Environment.Exit(exitCode: 0);
            }
        }

        void printSavedAs(string fileNameToSaveAs,
                          DateTime lastModified)
        {
            Console.WriteLine(value: "-- Saved as " +
                                     fileNameToSaveAs +
                                     " w/ TS " +
                                     lastModified.ToString(provider: CultureInfo.CurrentCulture));
        }
    }
}