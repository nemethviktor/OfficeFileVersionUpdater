using System;
using System.Globalization;
using System.IO;

namespace OfficeFileVersionUpdater;

public class OfficeHandler
{
    internal string FolderName;
    internal DateTime LastModified;

    protected DateTime GetLastModifiedDt(string fileToCheck)
    {
        return File.GetLastWriteTime(path: fileToCheck);
    }

    protected void PrintSavedAs(string fileNameToSaveAs,
        DateTime lastModified)
    {
        Console.WriteLine(value: "-- Saved as " +
                                 fileNameToSaveAs +
                                 " w/ TS " +
                                 lastModified.ToString(provider: CultureInfo.CurrentCulture));
    }
}