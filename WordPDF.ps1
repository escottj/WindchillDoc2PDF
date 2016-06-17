#############################################
# Original Script Created By: Erick Johnson #
# Created: April 30, 2016                   #
# Last Modified: June 16, 2016              #
# Version: 1.0                              #
# Supported Office: 2010, 2013, 2016        #
# Supported PowerShell: 4, 5                #
# Copyright © 2016 Erick Scott Johnson      #   
# All rights reserved.                      #
#############################################
#Convert arguments to variables
$Input  = $args[0]
$Output = $args[1]

#Create output file
New-Item $Output -Type file

#Read input file
$Array = @()
For ($i = 0;$i -le 5;$i++) {
    $Array += Get-Content $Input | ForEach {($_ -split '\s+',6)[$i]}
}
#$Array[0] = Filename
#$Array[1] = Extension
#$Array[2] = Rep Type (Ignore)
#$Array[3] = Not Used
#$Array[4] = Exported File Input Path
#$Array[5] = Path for PDF to go
$Path = $Array[4] + '\'
$Ext  = $Array[1]
$WCDTI = $Array[4] + '\wcdti.xml'
#Must get file this way instead of directly from array result 
#because of '@_' filename issue with Windchill
$File = Get-Item $Path* -Include *$Ext
$Filename = [System.IO.Path]::GetFileNameWithoutExtension($File)
$PDF  = $Array[5] + '\' + $Filename  + '.pdf'

#Check for desktop folder
If (-Not (Test-Path C:\Windows\SysWOW64\config\systemprofile\Desktop))
{
    New-Item C:\Windows\SysWOW64\config\systemprofile\Desktop -Type directory
}

#Open file
$Word = New-Object -ComObject Word.Application
$Doc  = $Word.Documents.Open([string]$File)

#Check version of Word installed and create PDF
$Version = $Word.Version
If ($Version -eq '16.0' -Or $Version -eq '15.0') {
    $Doc.SaveAs($PDF, 17) 
    $Doc.Close($False)  
}
ElseIf ($Version -eq '14.0') {
    $Doc.SaveAs([ref] $PDF,[ref] 17)
    $Doc.Close([ref]$False)
}

#Close Word
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
$Word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)

#Cleanup
Remove-Variable Word
$attHash.Clear()

#Tell Windchill PDF is ready
[IO.File]::WriteAllLines($Output, '0 ' + $PDF)