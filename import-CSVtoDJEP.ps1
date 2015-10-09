#DJEP CSV Music import
#Created By @JDogHerman

param (
    [string]$WindowTitle = "Music - Google Chrome",
    [string]$Path = $inputFile
    )


#csv file should be formatted with Classification,Song,Artist,Location,Comment,
#Any not needed field should be left blank

[System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null
add-type -AssemblyName microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms

$myDialog = New-Object System.Windows.Forms.OpenFileDialog
$myDialog.Title = “Please select the import file”

$myDialog.InitialDirectory = “%userprofile%”

$myDialog.Filter = “CSV Files (*.csv)|*.csv|All Files (*.*)|*.*”

$result = $myDialog.ShowDialog()

If($result -eq “OK”) {

$inputFile = $myDialog.FileName

# Continue working with file

}

else {

Write-Host “Cancelled by user”
exit
}


$csv = import-csv $inputFile

start-sleep -Milliseconds 500
[Microsoft.VisualBasic.Interaction]::AppActivate($WindowTitle)
start-sleep -Milliseconds 500
foreach ($line in $csv) {
[System.Windows.Forms.SendKeys]::SendWait($line.Classification)
start-sleep -Milliseconds 50
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
start-sleep -Milliseconds 50
[System.Windows.Forms.SendKeys]::SendWait($line.Song)
start-sleep -Milliseconds 50
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
start-sleep -Milliseconds 50
[System.Windows.Forms.SendKeys]::SendWait($line.Artist)
start-sleep -Milliseconds 50
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
start-sleep -Milliseconds 50
[System.Windows.Forms.SendKeys]::SendWait($line.Location)
start-sleep -Milliseconds 50
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
start-sleep -Milliseconds 50
[System.Windows.Forms.SendKeys]::SendWait($line.Comment)
start-sleep -Milliseconds 50
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
start-sleep -Milliseconds 50
}

Write-Host "Finished"
