# Legacy Office File Version Updater & TimeStamp Retainer

## What this does

The script takes legacy-format Office files (Word/Excel/PowerPoint) and converts them to new-format files. Read the "File Version support" block on what precisely happens and what it can do.

The added bonus here is that we retain the old "last-write-time" timestamp. In other words, if your "2003_budget.xls" had a timestamp of say `2003-01-01 12:27:00` then your new "2003_budget.xlsx" will also have the same TS.

Also just to be clear the process deletes the old files upon saving the new one, so in the case above, you'd just be left with one file, the xlsx. As always, make a backup and ensure you're happy with the outcome, then consider disposing of your backup.

So really the process is:
1. Scans the given folder(s) and collates relevant files into a list
2. Loops through the list and for each file:
3. Takes note of the last-modified TS
4. Attempts to open the file
5. Attempts to resave the file in modern format
6. Closes the file
7. Stamps the old TS on the new file
8. Deletes the original if Save was successful

## Usage

Download & unzip the whole contents of the 7z file from Releases (click Assets within, if you can't see it). 

Once unzipped, run the `OfficeFileVersionUpdater.exe` executable with the `-f` parameter for the folder you want to parse, i.e., `OfficeFileVersionUpdater.exe -f C:\something` or use double quotes if your folder has spaces in its name such as `OfficeFileVersionUpdater.exe -f "C:\something something"`. 

You can optionally add any of the `--SkipWord`, `--SkipExcel`, `--SkipPowerPoint` (not case-sensitive) parameters to instruct the code to skip the relevant app's files. E.g.: `OfficeFileVersionUpdater.exe -f C:\something --SkipWord --SkipExcel` will only parse PowerPoint files in the C:\something folder 
Parameters are not case sensitive and note that the "skips" have no shorthand abbreviations and so while `-f` is just one count of `-` the skips are two counts (`--`).

The script processes folders recursively so you only need to specify the root folder of your hierarchy.

## System Requirements

Although some DLL files are included in the release but you'll actually need a working version of Office 15+ (2013 that is. The relevant connector files between C# and Office are from the 2013 version, there aren't any newer ones as of 2023)

Windows 7 or newer is also required.

You'll obviously need to have access to your file passwords etc where applicable.

## What this does not (do)

Apart from the "what this does", nothing else. The app is not handling errors in the target apps related to things like "the document/workbook/presentation needs to be repaired etc." It's unlikely this would ever happen (see below for File Version support as to why).

I didn't add in the functionality to process non-Office files (particularly thinking very-legacy-non-Office here [MS Works etc].) -- if there is a need for this, open a ticket and I'll see what I can do.

Also, the app won't update things like `PivotTableVersion` (not explicitly anyway).

## File Version support

### TL;DR

I've tested as far back as files saved by Office 4.3 in Windows 3.11 and those worked *when explicitly allowed in Trust Center*. Older files might be supported if they can be opened in Word/Excel/PowerPoint natively.
The script parses .doc, .docx, .xls, .ppt and .pps files currently.

**For the "x-files..." (as in PPTX, XLSX, DOCX):**

- PowerPoint and Excel have no "compatibility mode". All PPTX/XLSX files are the same version. Due to this the code entirely ignores PPTX/PPTM/XLSX/XLSM files as there's no benefit to re-saving them.
- DOCX files do have "compatibility mode" and therefore they _are_ processed. 
- See note further below regarding pre-release Office 2007 formats.

**For the legacy file versions (non-x-files):**

- Opening very old (like pre-Office97) files is generally disabled by Office security settings. Pre-Office97 file interactions must be enabled in Word's/Excel's/PowerPoint's Trust Center.

### Way Too Long

I've been pointed out that the original readme wasn't clear enough about this so here's a much longer rundown of the situation. 
C# (the language this is written in) has native connectors, called interops. Interops enable the programmer to command the target app (e.g. Word) do to things it can natively do, basically enabling some level of automation and eliminating manual stuff - it's like Macros in Excel (or Word etc.) just written in a different language. These connectors don't add non-existing functionality to the native apps.

With that out of the way, the various file extensions (.doc, .xls...) aren't "obviously version-specific" (like there's no .do1, .do2 etc for version 1, version 2 files) and Windows doesn't know what version they are unless the user opens them in Word (etc.) and they either work, or they don't. Hypothetically [here's the very official list](https://learn.microsoft.com/en-us/deployoffice/compat/office-file-format-reference) of file formats supported by Office, but that list is incorrect (e.g. it says "Opening or saving to PowerPoint 95 (or earlier) file formats", which is bs because I've tried opening one and it worked).

**Word Files**

According to [Wikipedia](https://en.wikipedia.org/wiki/Microsoft_Word#Filename_extensions):

> Although the  [`.doc`](https://en.wikipedia.org/wiki/Doc_(computing)
> "Doc (computing)")  extension has been used in many different versions
> of Word, it actually encompasses four distinct file formats:
> 
> 1.  Word for DOS
> 2.  Word for Windows 1 and 2; Word 3 and 4 for Mac OS
> 3.  Word 6 and Word 95 for Windows; Word 6 for Mac OS
> 4.  Word 97 and later for Windows; Word 98 and later for Mac OS

As said above the first 3 of those 4 are generally blocked from opening in modern Word versions unless unlocked in the Trust Center.

**Excel Files**

For Excel documents [here's](http://www.openoffice.org/sc/excelfileformat.pdf) a very technical explanation (pages 10-13) but suffice to say that there were 5 different versions that were all just called .xls and again anything older than Office 97 (up until and including *Excel 5.0/7.0 workbook*) is blocked by default and needs to be unblocked.

**PowerPoint Files**

Again as per [Wikipedia](https://en.wikipedia.org/wiki/Microsoft_PowerPoint#Binary_%281987%E2%80%932007%29):

> Early versions of PowerPoint, from 1987 through 1995 (versions 1.0
> through 7.0), evolved through a sequence of binary file formats,
> different in each version, as functionality was
> added.[[258]](https://en.wikipedia.org/wiki/Microsoft_PowerPoint#cite_note-early-file-compatibility-258)
> This set of formats were never documented, but an open-source 
> _libmwaw_  (used by  [LibreOffice](https://en.wikipedia.org/wiki/LibreOffice
> "LibreOffice")) exists to read
> them.[[259]](https://en.wikipedia.org/wiki/Microsoft_PowerPoint#cite_note-259)
> 
> A stable binary format (called a .ppt file, like all earlier binary
> formats) that was shared as the default in PowerPoint 97 through
> PowerPoint 2003 for Windows...

I think if you're still reading this then by that point you know where to change the relevant setting (++ refer to LibreOffice above if really needed).

Really the point is, try to open a file in the normal front-end (Word etc) and see if it can be opened or not.

### Notes re Beta1 & Beta2 Office 2007 files

If you happen to have some files saved by pre-release versions of Office 2007 (in particular Beta 1 and 2) then this app won't solve your problems. One is that as per above the script ignores PPTX/XLSX files entirely but that aside Post-Beta2-Office2007 apps are not capable of reading Beta1 & Beta2 files because they are not the same format. So, the app will try to command Word to open the DOCX (because we can't get the versioning flag without opening a file) but will fail (because Word can't read it if your file was saved in _not_ B2TR. Excel can't read even the B2TR files.).

The best you can do is get a VM Windows 7 and a pre-release Office 2007 ISO from somewhere (message me) and then manually rebuild your files in a separate VM. 

## Further notes

I didn't actually test this on non `en-GB` machines, but I expect it should run ok.

If you have an Office 4.x or earlier file then the original files will be full-caps (like DOC1.DOC) and so the new file will save as "DOC1.DOCx" (x being low-cap). Soz.

Against my better visual judgement, I have coded so that the various apps show (rather than running entirely hidden in the background.) - this causes a lot of flickers but allows the user to interact with the apps should there be a need for it. Basically, due to the lack of proper commanding capabilities in the interops (connectors) this is to be deemed a lesser evil. Ugly but "works".

If you have VBA-enabled files in legacy formats there is a likelihood that they will break after having moved them to modern-format, espc if you have IntPTRs in 32bits realm and those are being moved into 64bits. This is "normal" and is unrelated to the script here. It's just Office being ...Office.

The usual about "I take no responsibilities and there are no warranties of any sorts" applies here.