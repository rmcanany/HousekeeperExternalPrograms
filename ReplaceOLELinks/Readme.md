# Replace OLE Links

## Description

A program to replace `SmartFrames2d` with `Blocks` in draft files.  A `SmartFrame2d` is the OLE-linked entity created with the `Insert>Object` command in Solid Edge.  

The program has been tested on exactly two files.  Don't even *think* about running it on production work without extensive testing on backups.

## Run Modes

The program can run stand-alone or in batch mode.  To run stand-alone, double-click `ReplaceOLELinks.exe`.  Solid Edge must be running with a draft file open.  

Batch mode uses Housekeeper's `Run External Program` command.  On the **Task Tab**, enable the command, then click the `Program` button and select the executable named above.

In either case, Windows may have placed a lock on the folder containing the executable.  Right-click it and make sure it is not `Blocked` or `Read-only`.  Also, the first time you run the program you may get a `Windows protected your PC` message.  You can click `More info` followed by `Run anyway` to launch it.

## Setup

A text file, `program_settings.txt`, holds the configuration options for the program.

### Report links only

This is to make a "dry run" of your files.  No changes are made.  Rather, it analyzes every file and reports the name of each `SmartFrame2d` it encounters.  

This is a way for the user to get an understanding of the file contents, which is intended to help configure the program correctly.

To enable, enter `True`, otherwise `False`.

`ReportLinksOnly = True`

### Background sheet names

Identifies the sheet names in the file to process.

- You can specify a single name: `BackgroundSheetNames = Hardigg D`

- A comma-delimited list: `BackgroundSheetNames = Hardigg C, Hardigg D`

- Or all sheets: `BackgroundSheetNames = *`

### Template filename

The name of the file that holds the `Blocks` that will replace the `SmartFrames2d`.  The full path is required.  For example:

`TemplateFilename = C:\data\CAD\TOLERANCE BLOCKS.dft`

### Link filenames

A comma-delimited list of `SmartFrame2d` names to replace with `Blocks`.  Only the file name, not the path, is needed.

`LinkFilenames = PETOLERANCE.doc, , PUTOLERANCE.doc`

You may notice that the second item on the list does not have a name.  That is how a `SmartFrame2d` without an external link is reported by Solid Edge.  It is the main reason the `ReportLinksOnly` option is included.

### Block names

A comma-delimited list of the `Blocks` to replace the `SmartFrames2d`.  Its order matches that of LinkFileNames.  The first block will replace the first link, the second block replaces the second link, etc.

`BlockNames = PE TOLERANCE, PU TOLERANCE, PU TOLERANCE`
