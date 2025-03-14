﻿$zDriveDestDir = "Z:\GMEP Engineers\Users\GMEP Softwares\AutoCAD Commands\GMEPElectricalCommands"
$date = Get-Date -Format "yyyy-MM-dd HH-mm-ss"
 mkdir -Force -Path "$zDriveDestDir\$date"
 mkdir -Force -Path "$zDriveDestDir\latest"
Copy-Item -Force -Recurse -Path "bin\Debug\*" -Destination "$zDriveDestDir\latest"
Copy-Item -Force -Recurse -Path "bin\Debug\*" -Destination "$zDriveDestDir\$date"
