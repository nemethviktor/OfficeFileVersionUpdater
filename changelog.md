# OfficeFileVersionUpdater Changelog

**v1.0.5 [20261025]**
- NEW & UPDATED:
  - Build output (`bin\` and `obj\`) now goes to `%LOCALAPPDATA%\build\OfficeFileVersionUpdater\` via a new `Directory.Build.props`, so rebuilding no longer churns the synced repository folder.
  - File collation now walks the folder tree once instead of up to three times, and skips Office `~$` lock files up front.
  - Added the application icon.
  - (The app never shelled out to PowerShell, so there was nothing to replace with native C# calls; it already uses the Office interops and the registry API directly.)
- BUGS & FIXES:
  - Upper-case extensions (e.g. `DOC1.DOC`, `BUDGET.XLS`, `SLIDES.PPT`) are now handled correctly. Previously Excel/PowerPoint files crashed the run with a NullReferenceException, and `.DOC` files were re-saved under the old name instead of converted to `.docx`.
  - A file that fails to open no longer crashes the run (`ProcessAndSaveFile` was handed a null document).
  - A failed save in Word no longer throws on a null file name, and the document is closed after a failed save in all three apps.
  - The Office apps are now always quit (`try/finally`), so an unexpected error no longer leaves hidden Word/Excel/PowerPoint processes running.
  - PowerPoint failing to start reported "Excel installation not found"; it now reports PowerPoint.

**v1.0.4 [20230807]**
- NEW & UPDATED:
  - Extended the readme by quite a bit
  - Added options for skipping the various Applications
- BUGS & FIXES:
  - N/A


**v1.0.3 [20230805]**
- NEW & UPDATED:
  - Rewrote the code to be more OOP. 
- BUGS & FIXES:
  - N/A

**v1.0.2 [20230805]**
- NEW & UPDATED:
  - Included the Office DLL files. You'll still need Office installed to run but this adds some level of compatibility.
  - Reworked some of the error/exit handling
  - Removed unused references (code stuff)
  - Reformatted the code
- BUGS & FIXES:
  - N/A

**v1.0.1 [20230804]**
- NEW & UPDATED:
  - Changed the command-line argument handling logic. This is largely so that the introduction of further (future) parameters becomes a lot easier to handle.
  - Updated the readme
- BUGS & FIXES:
  - N/A

**v1.0.0 [20230804]**
- NEW & UPDATED:
  - Initial release
- BUGS & FIXES:
  - N/A