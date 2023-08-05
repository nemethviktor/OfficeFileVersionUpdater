using System;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using static OfficeFileVersionUpdater.Program;

namespace OfficeFileVersionUpdater;

public class PowerPointHandler : OfficeHandler
{
    private Application _powerPointApp;


    /// <summary>
    ///     Fetches a new PowerPoint process
    /// </summary>
    /// <returns>A PowerPoint process</returns>
    internal Application StartApp()
    {
        try
        {
            _powerPointApp = new Application
            {
                WindowState = PpWindowState.ppWindowMinimized,
                Visible = MsoTriState.msoTrue
            };
        }
        catch
        {
            ProgramHelpers.ExitWithMessage(exitReason: ExitReasons.ExcelNotInstalled);
        }

        return _powerPointApp;
    }

    internal void ProcessAndSaveFile(Presentation powerPointPres)
    {
        string origFileName = Path.Combine(path1: FolderName, path2: powerPointPres.Name);
        string newFileName = null;

        bool savedOk = false;
        try
        {
            newFileName = SaveActualFile(presentation: powerPointPres,
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

    public Presentation OpenFile(string fileNameWithPath, Application powerPointApp)
    {
        FileInfo fi = new(fileName: fileNameWithPath);
        Presentation powerPointPres = null;
        if (!fi.Name.Contains(value: "~$") && (fi.Extension == ".ppt" || fi.Extension == ".pps"))
            try
            {
                Console.WriteLine(value: "Processing " + fileNameWithPath);
                FolderName = fi.DirectoryName;
                LastModified = GetLastModifiedDt(fileToCheck: fileNameWithPath);
                powerPointPres = powerPointApp.Presentations.Open(FileName: fileNameWithPath,
                    ReadOnly: MsoTriState.msoTrue,
                    Untitled: MsoTriState.msoFalse, WithWindow: MsoTriState.msoFalse);
            }
            catch (Exception ex)
                // System.Runtime.InteropServices.COMException -- generally this will be stuff like pre-Office97 files. 
                // HRESULT: 0x800A03EC -- you most likely provided the wrong password.
            {
                Console.WriteLine(value: "-- " + ex.Message);
            }

        return powerPointPres;
    }

    private string SaveActualFile(Presentation presentation, out bool savedOk)
    {
        string newFileName = Path.Combine(path1: FolderName,
            path2: presentation.HasVBProject ? presentation.Name + "m" : presentation.Name + "x");
        FileInfo fi = new(fileName: Path.Combine(path1: FolderName, path2: presentation.Name));
        if (presentation.HasVBProject)
            presentation.SaveAs(FileName: newFileName,
                FileFormat: fi.Extension.ToLower() ==
                            ".ppt"
                    ? PpSaveAsFileType.ppSaveAsOpenXMLPresentationMacroEnabled
                    : PpSaveAsFileType.ppSaveAsOpenXMLShowMacroEnabled);
        else
            presentation.SaveAs(FileName: newFileName,
                FileFormat: fi.Extension.ToLower() ==
                            ".ppt"
                    ? PpSaveAsFileType.ppSaveAsOpenXMLPresentation
                    : PpSaveAsFileType.ppSaveAsOpenXMLShow);

        savedOk = true;
        presentation.Close();
        return newFileName;
    }

    public void QuitApp()
    {
        _powerPointApp.Quit();
        Console.WriteLine(value: "PowerPoint files done");
    }
}