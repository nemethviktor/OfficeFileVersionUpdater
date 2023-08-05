using System;
using System.Collections.Generic;
using Microsoft.Win32;
using static OfficeFileVersionUpdater.Program;

internal static class ProgramHelpers
{
    /// <summary>
    ///     Checks if a registry subkey exists
    /// </summary>
    /// <param name="hiveHklmOrHkcu"></param>
    /// <param name="registryRoot"></param>
    /// <returns></returns>
    /// <exception cref="System.InvalidOperationException"></exception>
    internal static bool RegistrySubKeyExists(string hiveHklmOrHkcu,
        string registryRoot)
    {
        RegistryKey root = hiveHklmOrHkcu.ToUpper() switch
        {
            "HKLM" => Registry.LocalMachine.OpenSubKey(name: registryRoot, writable: false),
            "HKCU" => Registry.CurrentUser.OpenSubKey(name: registryRoot, writable: false),
            _ => throw new InvalidOperationException(
                message: "parameter registryRoot must be either \"HKLM\" or \"HKCU\"")
        };
        bool valExists = root != null;
        return valExists;
    }

    /// <summary>
    ///     Attempts to get the max version of Office installed based on available registry keys
    /// </summary>
    /// <returns></returns>
    internal static string GetOfficeVer()
    {
        List<string>
            officeVersionList =
                new(collection: new[]
                    { "17.0", "16.0", "15.0", "14.0", "12.0" }); // no 17 as of 2023 but it'm "future-proofing".

        foreach (string version in officeVersionList)
            if (RegistrySubKeyExists(hiveHklmOrHkcu: "HKLM", registryRoot: "SOFTWARE\\Microsoft\\Office\\" + version))
                return version;

        return null;
    }

    /// <summary>
    ///     Attempts to clear the "blacklist" of problematic files for each app.
    /// </summary>
    /// <param name="whichApp"></param>
    /// <param name="officeVer"></param>
    internal static void ClearBlackList(string whichApp,
        string officeVer)
    {
        // clear registry key for blacklisted files
        // basically idgaf
        if (officeVer != null)
        {
            RegistryKey blacklistKey =
                Registry.CurrentUser.OpenSubKey(
                    name: "Software\\Microsoft\\Office\\" + officeVer + "\\" + whichApp + "\\Resiliency",
                    writable: true);
            if (blacklistKey != null)
                try
                {
                    blacklistKey.DeleteSubKeyTree(subkey: "DisabledItems", throwOnMissingSubKey: false);
                }
                catch
                {
                    // ignored
                }
        }
    }

    /// <summary>
    ///     Exits the program with the given code. Easier to manage in one spot tbh.
    /// </summary>
    /// <param name="exitReason"></param>
    internal static void ExitWithMessage(ExitReasons exitReason)
    {
        switch (exitReason)
        {
            case ExitReasons.Ok:
                Console.WriteLine(value: "All seems to be well.");
                break;
            case ExitReasons.NoFolderPassed:
                Console.WriteLine(value: "No folder provided.");
                break;
            case ExitReasons.InvalidFolder:
                Console.WriteLine(value: "The path doesn't exist.");
                break;
            case ExitReasons.WordNotInstalled:
                Console.WriteLine(value: "Word installation not found.");
                break;
            case ExitReasons.ExcelNotInstalled:
                Console.WriteLine(value: "Excel installation not found.");
                break;
            case ExitReasons.PowerpointNotInstalled:
                Console.WriteLine(value: "Powerpoint installation not found.");
                break;
        }

        Console.WriteLine(value: "Exiting.");
        Environment.Exit(exitCode: (int)exitReason);
    }
}