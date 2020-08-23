#Created by Kyle Ericson
#Note the C:\IntuneWinAppUtil.exe file needs to be in the root of C:\

Remove-Item $env:TEMP -Force -Recurse -ErrorAction SilentlyContinue

Function Get-FileName($initialDirectory)
{
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Please Select File"
    $OpenFileDialog.InitialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    # Out-Null supresses the "OK" after selecting the file.
    $OpenFileDialog.ShowDialog() | Out-Null
    $Global:SelectedFile = $OpenFileDialog.FileName
    $Global:SelectedFilePath = split-path -path $SelectedFile
    $Global:SelectedFileName = [System.IO.Path]::GetFileName("$SelectedFile")
}

Get-FileName
echo $SelectedFile
echo $SelectedFilePath
echo $SelectedFileName 

Start-Process C:\IntuneWinAppUtil.exe "-c $SelectedFilePath -s $SelectedFileName -o $SelectedFilePath" -Wait 
