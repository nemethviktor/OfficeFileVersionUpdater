using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using static OfficeFileVersionUpdater.Program;

namespace OfficeFileVersionUpdater;

public class ExcelHandler : OfficeHandler
{
    private Application _excelApp;

    /// <summary>
    ///     Fetches a new Excel process
    /// </summary>
    /// <returns>A Excel process</returns>
    internal Application StartApp()
    {
        try
        {
            _excelApp = new Application
            {
                WindowState = XlWindowState.xlMinimized,
                AskToUpdateLinks = false,
                Visible = true
            };
        }
        catch
        {
            ProgramHelpers.ExitWithMessage(exitReason: ExitReasons.ExcelNotInstalled);
        }

        return _excelApp;
    }

    internal void ProcessAndSaveFile(Workbook excelWbk)
    {
        string origFileName = Path.Combine(path1: FolderName, path2: excelWbk.Name);
        string newFileName = null;
        // check compatibility [aka extension-eqsue]
        if (excelWbk.Excel8CompatibilityMode)
        {
            bool savedOk = false;
            try
            {
                newFileName = SaveActualFile(workbook: excelWbk,
                    savedOk: out savedOk);
            }
            catch
            {
                Console.WriteLine(value: "-- Save failed for " + origFileName);
            }

            if (savedOk)
            {
                // kill old version file if applicable
                File.Delete(path: origFileName);
                File.SetLastWriteTime(path: newFileName, lastWriteTime: LastModified);
                PrintSavedAs(fileNameToSaveAs: newFileName,
                    lastModified:
                    LastModified); // kinda fake here for lastModified but if we got this far then setting the TS should be ok.
            }
        }
        else // current-version file.
        {
            excelWbk.Close(SaveChanges: false);
            Console.WriteLine(value: "- Ignored up-to-date version file: " + origFileName);
        }
    }

    public Workbook OpenFile(string fileNameWithPath, Application excelApp)
    {
        FileInfo fi = new(fileName: fileNameWithPath);
        Workbook excelWbk = null;
        if (!fi.Name.Contains(value: "~$") && fi.Extension == ".xls")
            try
            {
                Console.WriteLine(value: "Processing " + fileNameWithPath);
                FolderName = fi.DirectoryName;
                LastModified = GetLastModifiedDt(fileToCheck: fileNameWithPath);
                excelWbk = excelApp.Workbooks.Open(Filename: fileNameWithPath);
            }
            catch (Exception ex)
                // System.Runtime.InteropServices.COMException -- generally this will be stuff like pre-Office97 files. 
                // HRESULT: 0x800A03EC -- you most likely provided the wrong password.
            {
                Console.WriteLine(value: "-- " + ex.Message);
            }

        return excelWbk;
    }

    private string SaveActualFile(Workbook workbook, out bool savedOk)
    {
        string newFileName =
            Path.Combine(path1: FolderName, path2: workbook.HasVBProject ? workbook.Name + "m" : workbook.Name + "x");
        workbook.SaveAs(Filename: newFileName,
            FileFormat: workbook.HasVBProject
                ? XlFileFormat.xlOpenXMLWorkbookMacroEnabled
                : XlFileFormat.xlOpenXMLWorkbook);
        savedOk = true;
        workbook.Close(SaveChanges: false);
        return newFileName;
    }

    public void QuitApp()
    {
        _excelApp.Quit();
        Console.WriteLine(value: "Excel files done");
    }
}