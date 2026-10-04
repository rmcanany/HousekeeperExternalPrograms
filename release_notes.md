<div class="center">
  <p align=center>
  <img src="logo.png" width=50%;>
  <p align=center>
  <span class="description">Robert McAnany 2026</span>
</div>

# Release Notes

Companion repo with instructions and program examples for the 
Solid Edge Housekeeper `Run external program` task.

## V2026.2

Added a new external program, `CheckDisplayConfigurations`.  Added code snippets `FaceStylesExport`, `FaceStylesImport`, `ModifyFaceStyle`, `UnitsToANSI`, and `UnitsToISO`.  

Updated the code snippet documentation to better describe how to find and navigate the API to help build your own.

Changed all external program parameter file names to `program_setting.txt`, mostly to simplify the newly-automated Release creation script.

## V2026.1

Added `ReplaceOLELinks`, `UpdatePartStyleFromTemplate` and several Snippet examples.

## V2025.3

Added a code snippet contributed by @ih0nza `SaveAndTogglePreviewGeometry.snp`.  Thank you!

## V2025.2

Added `Snippets`, `ThinPartToSheetmetal` and `CreateFlatPattern`.

## V2025.1

Updated error handling and configuration input.

## V2024.0

Added example programs `CompareFlatAndModelVolumes`, `QtyFromAssy`.
Updated `AddRemoveCustomProperties`.

## V2023.1

Changed error handling to be more consistent with Housekeeper.
See the [**Readme**](Readme.md) for details.

Added `ChangeToInchAndSaveAsFlatDXF.vbs`,
a more complete VBScript example.
With VBScript you only need a text editor to write a
program.  No need to install or learn Visual Studio!

## v 0.1.3

Added BreakExcelLinks and AssemblyReport.  These two macros illustrate 
the use of VBScript with Housekeeper.  

Modified FitIsoView, adding the function 'GetConfiguration()'.  This 
reads in Housekeeper's defaults.txt, making selected options available 
to the programmer.

## v 0.1.2

Modified RenameSheets

## v 0.1.1

Added AddRemoveCustomProperties

## v 0.1.0

Initial release