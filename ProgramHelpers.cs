using System;
using System.Collections.Generic;
using Microsoft.Win32;

internal static class ProgramHelpers
{
    /// <summary>
    ///     Checks if a registry subkey exists
    /// </summary>
    /// <param name="hive_HKLM_or_HKCU"></param>
    /// <param name="registryRoot"></param>
    /// <returns></returns>
    /// <exception cref="System.InvalidOperationException"></exception>
    internal static bool RegistrySubKeyExists(string hive_HKLM_or_HKCU,
                                              string registryRoot)
    {
        RegistryKey root = hive_HKLM_or_HKCU.ToUpper() switch
        {
            "HKLM" => Registry.LocalMachine.OpenSubKey(name: registryRoot, writable: false),
            "HKCU" => Registry.CurrentUser.OpenSubKey(name: registryRoot, writable: false),
            _ => throw new InvalidOperationException(message: "parameter registryRoot must be either \"HKLM\" or \"HKCU\"")
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
                new(collection: new[] { "17.0", "16.0", "15.0", "14.0", "12.0" }); // no 17 as of 2023 but it'm "future-proofing".

        foreach (string version in officeVersionList)
        {
            if (RegistrySubKeyExists(hive_HKLM_or_HKCU: "HKLM", registryRoot: "SOFTWARE\\Microsoft\\Office\\" + version))
            {
                return version;
            }
        }

        return null;
    }

    /// <summary>
    ///     Attempts to clear the "blacklist" of problematic files for each app.
    /// </summary>
    /// <param name="whichApp"></param>
    /// <param name="officeVer"></param>
    internal static void clearBlackList(string whichApp,
                                        string officeVer)
    {
        // clear registry key for blacklisted files
        // basically idgaf
        if (officeVer != null)
        {
            RegistryKey blacklistKey =
                Registry.CurrentUser.OpenSubKey(
                    name: "Software\\Microsoft\\Office\\" + officeVer + "\\" + whichApp + "\\Resiliency", writable: true);
            if (blacklistKey != null)
            {
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
    }
}