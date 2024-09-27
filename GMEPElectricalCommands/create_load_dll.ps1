$currentDirectory = (Get-Item .).FullName -replace "\\", "/"
Write-Output "(command `"netload`" `"$currentDirectory/bin/Debug/ElectricalCommands.dll`")" | Out-File "load_dll.lsp"