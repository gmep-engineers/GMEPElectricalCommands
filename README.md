### Development Environment Setup
1. Download [Visual Studio](https://visualstudio.microsoft.com/) installer
1. During install, select **.NET desktop development**
1. Open Visual Studio
1. Select *Clone a Repository*
1. Clone this repository
1. Open the project in Windows Explorer
1. In the `GMEPElectricalCommands` folder, right-click `create_load_dll.ps1` and click *Run with PowerShell*.
1. Open blank file in AutoCAD and run the `APPLOAD` command.
1. Under *Startup Suite*, click *Contents*.
1. Click *Add*.
1. Navigate to the same `GMEPElectricalCommands` folder above. Select `load_dll.lsp`. Click *Open*.
1. Click *Close* to close the Startup Suite window.
1. Click *Close* to close the Load/Unload Applications window.
1. Close and reopen AutoCAD. Click *Always Load* at the security prompts.

## If Visual Studio doesn't have an option to build
Src: https://stackoverflow.com/questions/75132749/project-loading-error-due-to-missing-defineconstants-value-but-the-values-are-d
1. Close Visual Studio
1. Open `GMEPElectricalCommands.sln` in a text editor
1. Replace all occurrences of the `9A19103F-16F7-4668-BE54-9A1E7A4F7556` GUID with `FAE04EC0-301F-11D3-BF4B-00C04F79EFBC`
1. Save
1. Reopen the solution in Visual Studio

## If the solution doesn't build
1. Try moving the `packages` folder one folder up such that it's in `GMEPElectricalCommands`, not `GMEPElectricalCommands/GMEPElectricalCommands`

You can now build and test the GMEPElectricalCommands project. AutoCAD should open each time a build occurs. AutoCAD must be closed in order to commit changes or edit `[Design]` components. 