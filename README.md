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

You can now build and test the GMEPElectricalCommands project. AutoCAD should open each time a build occurs. AutoCAD must be closed in order to commit changes or edit `[Design]` components. 