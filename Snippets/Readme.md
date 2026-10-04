# Code Snippet Examples

Some sample code to get you started with Housekeeper's `Run External Program` command `Code Snippet` option. 

The samples probably won't be exactly what you need.  You can learn more about what is possible with the API documentation.  To find it, in Solid Edge, go to the `File` menu > `Discover` pane > `Learn` tab > `My Links` section > `Programming with Solid Edge` button.  The `Type Libraries` section is the most useful.

For example, say this code is almost what you want, except instead of changing the shininess, you want to replace the texture image.

```
    For Each FaceStyle As SolidEdgeFramework.FaceStyle In FaceStyles
        If FaceStyle.StyleName = "Aluminum" Then
            FaceStyle.Shininess = 0.75
            Exit For
        End If
    Next
```

The code tells you where to look.  A `FaceStyle` is defined in the `SolidEdgeFramework` type library.  Open that and expand its `Objects` collection.  Scroll down to `FaceStyle` and open its `Members` list.  In there you will find `TextureFileName`.

By the way, the snippets are just text files.  You can edit them in Notepad.  What follows is a short description of what each example does.

## CountFlatPatternModelsAndFlatPatterns

Shows how to access flat patterns and the models from which they are derived.

## DeleteAllBlocks

Example of working with the blocks collection.

## EllipticalArcPerimeter

Illustrates how to work with drawing view geometry.  Calculates arc length via numerical integration.

## FaceStylesImport/FaceStylesExport

Shows how to read/write tab-separated-value files.  Used in this case to batch-update face styles.

## FitIso

Sample code showing the use of the Solid Edge StartCommand.  For links to available commands, see `interactive_edit_commands.txt` in your Housekeeper Preferences directory.

## IsWeldment

Demonstrates how to find if an assembly has been marked as a weldment, and populate a file property to make it accessible for external use.

## ModifyFaceStyle

Shows how to loop through the face style collection to find a certain one, then modify it.

## PartCopyFilenameToProperty

Similar to `IsWeldment`, in that is shows how to expose an internal value as a file property.

## SaveAndTogglePreviewGeometry

Contributed code that shows how to work with Solid Edge global parameters.

## SEAppHeightWidth

Simple example that demonstrates how to work with the Solid Edge window object.

## SendKeystrokes

Shows a technique to interact with a dialog box programmatically.  Can be handy when a function is not exposed in the API.  Can also be flaky, as it depends on window focus to operate correctly.

## SerialIncrementPropertyNumber

Shows how to save mutiple copies of a file with an increasing index number in the file name.

## ThinPartToSheetmetal

User request that converts uniform-thickness parts to sheetmetal.  Handy for imported step files.

## UnitsToANSI/UnitsToISO

Demonstrates how to work with the units system.

## UpdateOnSave

Shows how to enable the `Update physical properties on save` toggle.

## UpdatePartsList

Example showing how to work with drawing parts lists.

## UpdatePartsListStyleFromTemplate

Illustrates a method for working with external text files.

## UpdateThreadProperties

Contributed code that shows how to access threads via the `HoleData` object.
