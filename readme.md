# Legacy Office File Version Updater & TimeStamp Retainer

## What this does

The script takes legacy-format Office files (Word/Excel/PowerPoint) and converts them to new-format files.
The added bonus here is that we retain the old "last-write-time" timestamp. In other words, if your "2003_budget.xls" had a timestamp of idk 2003-01-01 12:27:00 then your new "2003_budget.xlsx" will also have the same TS.

### Notes re Office (file) versions

For the "x-files" (as in PPTX, XLSX, DOCX):

- PowerPoint and Excel have no "compatibility mode". All PPTX/XLSX files are the same version. Due to this the code entirely ignores PPTX/PPTM/XLSX/XLSM files as there's no benefit to re-saving them.
- DOCX files do have “compatibility mode" and therefore they _are_ processed. 
- See note below regarding pre-release Office 2007 formats.

For the legacy file versions (non-x-files):

- Opening very old (like pre-Office97) files is generally disabled by Office security settings. This can be amended on the user's computer in the independent apps' Trust Centers. I tried w/ Office 95 and 4.x file versions, and they worked. I presume that should do the trick.
- Joyfully enough the error messages returned are different for each app for the above issue. E.g., Word will actually report that the file is too old of a version and so you need to update settings in the Trust Center. PowerPoint on the other hand will just say "unable to open the file". Basically, if you suspect that the reason may be that the file is super old then adjust the settings accordingly.
- The oldest Office (file version) I have and can interact with is v4.3. That works, I've tested, subject to settings above and if you have even older files then you're in the 80s era...which...should hypothetically still work. If not, open a ticket.

## What this does not (do)

It doesn't do anything else. It's not currently handling errors in the target apps related to things like "the document/workbook/presentation needs to be repaired etc." It's unlikely this would ever happen (see generic explanation re: the Office app you have defines the capabilities).
I didn't add in the functionality to process non-Office files (particularly thinking very-legacy-non-Office here.) -- if there is a need for this, open a ticket.
Also, the app won't update things like PivotTableVersions (not explicitly anyway).

### Notes re Beta1 & Beta2 Office 2007 files

If you happen to have some files saved by pre-release versions of Office 2007 (in particular Beta 1 and 2) then this app won't solve your problems. One is that as per above the script ignores PPTX/XLSX files entirely but that aside Post-Beta2-Office2007 apps are not capable of reading Beta1 & Beta2 files because they are not the same format. So, the app will try to command Word to open the DOCX (because we can't get the versioning flag without opening a file) but will fail (because Word can't read it if your file was saved in _not_ B2TR. Excel can't read even the B2TR files.).
The best you can do is get a VM Windows 7 and a pre-release Office 2007 ISO from somewhere (message me) and then manually rebuild your files in a separate VM. 

## What you need

You'll need a working version of Office 15+ (2013 that is. The relevant connector files between C# and Office are from the 2013 version, there aren't any newer ones as of 2023)
You'll obviously need to have access to your file passwords etc if they are relevant.

## How to run

Run the executable with the first parameter being the folder you want to parse, i.e., `OfficeFileVersionUpdater.exe C:\something` or use double quotes if your folder has spaces in its name such as `OfficeFileVersionUpdater.exe "C:\something something"`. 
The script processes folders recursively so you are expected to have just one parameter.

## Further notes

I didn't actually test this on non EN-GB machines, but I should expect it should run ok.
This is for people with OCD but if you have an Office 4.x or earlier file then the original files will be full-caps (like DOC1.DOC) and so the new file will save as "DOC1.DOCx" (x being low-cap). Soz.
Against my better visual judgement, I have coded so that the various apps show (rather than running in the background.) - this causes a lot of flickers but allows the user to interact with the apps should there be a need for it. Basically, due to the lack of proper commanding capabilities in the interops (connectors) this is to be deemed a lesser evil. Ugly but "works".
If you have VBA-enabled files in legacy formats there is a likelihood that they will break after having moved them to modern-format, espc if you have IntPTRs in 32bits realm and those are being moved into 64bits. This is "normal" and is unrelated to the script here. It's just Office being ...Office.
The usual about "I take no responsibilities and there are no warranties of any sorts" applies here.