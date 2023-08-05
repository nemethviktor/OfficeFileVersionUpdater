using System;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace OfficeFileVersionUpdater;

public class WordHandler : OfficeHandler
{
    private Application _wordApp;


    /// <summary>
    ///     Fetches a new Word process
    /// </summary>
    /// <returns>A Word process</returns>
    internal Application StartApp()
    {
        try
        {
            _wordApp = new Application
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
        catch
        {
            ProgramHelpers.ExitWithMessage(exitReason: Program.ExitReasons.WordNotInstalled);
        }

        return _wordApp;
    }

    internal void ProcessAndSaveFile(Document wordDoc)
    {
        // check compatibility [aka extension-eqsue]
        // sub-15 is basically "old" refer to https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdcompatibilitymode?view=word-pia
        string origFileName = Path.Combine(path1: FolderName, path2: wordDoc.Name);
        string newFileName = null;
        if (wordDoc.CompatibilityMode < 15)
        {
            if (wordDoc.Name.EndsWith(value: ".doc")) // super old
            {
                bool savedOk = false;
                try
                {
                    newFileName = SaveActualFile(wordDoc: wordDoc, retainFileName: false,
                        savedOk: out savedOk);

                    PrintSavedAs(fileNameToSaveAs: newFileName,
                        lastModified:
                        LastModified); // kinda fake here for lastModified but if we got this far then setting the TS should be ok.
                }
                catch
                {
                    Console.WriteLine(value: "-- Save failed for " + origFileName);
                }

                if (savedOk)
                    // kill old version file if applicable
                    File.Delete(path: origFileName);
            }
            else // basically this is stuff like version 2007 and generally "early" x-format files.
            {
                newFileName = SaveActualFile(wordDoc: wordDoc, retainFileName: true,
                    savedOk: out _);
                Console.WriteLine(value: "-- Re-saved file as version 2016 (current).");
            }

            File.SetLastWriteTime(path: newFileName, lastWriteTime: LastModified);
        }
        // non-legacy // compatibility-mode file
        else
        {
            wordDoc.Close(SaveChanges: false);
            Console.WriteLine(value: "- Ignored up-to-date version file: " + origFileName);
        }
    }

    internal Document OpenFile(string fileNameWithPath, Application wordApp)
    {
        FileInfo fi = new(fileName: fileNameWithPath);
        Document wordDoc = null;
        if (!fi.Name.Contains(value: "~$") && fi.Extension.Contains(value: ".doc"))
            try
            {
                Console.WriteLine(value: "Processing " + fileNameWithPath);
                FolderName = fi.DirectoryName;
                LastModified = GetLastModifiedDt(fileToCheck: fileNameWithPath);
                wordDoc = wordApp.Documents.Open(FileName: fileNameWithPath);
            }
            catch (Exception ex)
                // System.Runtime.InteropServices.COMException -- generally this will be stuff like pre-Office97 files. 
                // HRESULT: 0x800A03EC -- you most likely provided the wrong password.
            {
                Console.WriteLine(value: "-- " + ex.Message);
            }

        return wordDoc;
    }

    private string SaveActualFile(Document wordDoc, bool retainFileName, out bool savedOk)
    {
        string newFileName = Path.Combine(path1: FolderName, path2: !retainFileName
            ? wordDoc.HasVBProject ? wordDoc.Name + "m" : wordDoc.Name + "x"
            : wordDoc.Name);

        wordDoc.SaveAs2(FileName: newFileName,
            FileFormat: wordDoc.HasVBProject
                ? WdSaveFormat.wdFormatXMLDocumentMacroEnabled
                : WdSaveFormat.wdFormatXMLDocument);

        // word saves in compatibility mode for some reason by default.
        wordDoc.Convert(); // update to current
        wordDoc.Save();
        wordDoc.Close();
        savedOk = true;
        return newFileName;
    }

    internal void QuitApp()
    {
        _wordApp.Quit(SaveChanges: false);
        Console.WriteLine(value: "Word files done");
    }
}