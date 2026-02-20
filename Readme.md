![Logo](logo.png)

# Solid Edge Housekeeper External Programs
2026 Robert McAnany

Description and examples for Solid Edge Housekeeper's `Run External Program` task.

The external program is a Console App that works on a single Solid Edge file.  Housekeeper serves up files one at a time for processing.  

Most developers will want the source code, however compiled versions are available as [<ins>**Releases**</ins>](https://github.com/rmcanany/HousekeeperExternalPrograms/releases/).

## Requirements

To work with Housekeeper's error reporting, the program should return an integer exit code, with `0` meaning success.  Anything else indicates an error.  If an exit code is not issued, Housekeeper assumes success.

In order to return an exit code, the declaration for `Main()` should be: `Function Main() As Integer`. The usual declaration, `Sub Main()`, does not return a value.

You can provide feedback to the user through Housekeeper's error reporting mechanism.  See details on how to do so in the example programs `CompareFlatAndModelVolumes` and `ChangeToInchAndSaveAsFlatDXF.vbs` for examples for VB.NET and VBScript, respectively.  

If the `ExitCode` is not `0`, and your program does not implement error reporting, Housekeeper simply lists the exit code.

Housekeeper launches the program as follows:

    Dim P As New Process
    Dim ExitCode As Integer
    Extension = IO.Path.GetExtension(Me.ExternalProgram)

    If Extension = ".ps1" Then
        P.StartInfo.FileName = "powershell.exe"
        P.StartInfo.Arguments = String.Format("-command {1}{0}{1}", Me.ExternalProgram, Chr(34))
        P.Start()
    Else
        P = Diagnostics.Process.Start(Me.ExternalProgram)
    End If

    P.WaitForExit()
    ExitCode = P.ExitCode


No arguments are passed to the program. If you need to get a value from Housekeeper, such as a template file location, you can use the function `GetSettings()` as shown in `FitIsoView`. The function parses the `form_main_settings.json` file passed into the macro's startup directory. The file is updated just before processing is launched. It should always reflect the current status of the form.

Housekeeper maintains a `reference` to the file being processed. If that reference is broken, an exception will occur. To avoid that, do not perform `Close()` or `SaveAs()` (`SaveCopyAs()` is OK) on the document.

No assumptions are made about what the external program does. If you change a file and want to save it, that needs to be in the program.  If you open another file, you need to close it. 

## Other repos and sources for macros

Sample programs may be the easiest way to get started automating Solid Edge yourself. Here are a few sites to check out:

Jason Newell: [**SolidEdgeCommunity**](https://github.com/SolidEdgeCommunity)

Tushar Suradkar: [**SurfAndCode**](http://www.surfandcode.in/2014/01/index-of-all-tutorials-on-this-solid.html)

Jason Titcomb: [**LMGiJason**](https://github.com/LMGiJason)

Wolfgang Hackl: [**Cadcam-Consult**](http://cadcam-consult.com/Page_00/index.html)

Kabir Costa: [**KabirCosta**](https://github.com/kabircosta)

Alan Pope: [**Cadcentral NZ**](https://www.cadcentral.co.nz/macros)

UK Dave: [**uk-dave**](https://github.com/uk-dave/SolidEdge)

Francesco Arfilli: [**farfilli**](https://github.com/farfilli)

Chris Clems: [**ChrisClems**](https://github.com/ChrisClems)

Derek Gooding: [**DerekGooding**](https://github.com/DerekGooding)

NosyBottle: [**NosyBottle**](https://github.com/Nosybottle)

ZaPpInG: [**lrmoreno007**](https://github.com/lrmoreno007)


To suggest other sources, including your own, feel free to message me, RobertMcAnany, on the [**Solid Edge Forum**](https://community.sw.siemens.com/s/topic/0TO4O000000MihiWAC/solid-edge).

If you're looking for someone to create a program *for* you, I know a couple of folks who do it for a living.  Message me on the Forum and I'll pass along the request.




