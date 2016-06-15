#############################################
# PowerShell PDF Generation Script for Word #
# © 2016 Erick Scott Johnson                #
# Working Word Versions: 2010, 2013, 2016   #
# Working PowerShell Versions: 4, 5         #
# Working File Formats: all Word readable   #
# Last Modified: June 5, 2016               #
# Version: 1.0                              #
#############################################
#Relevant Files
$File  = $args[0]

#Generate PDF Filename
$Path     = [System.IO.Path]::GetDirectoryName($File)
$Filename = [System.IO.Path]::GetFileNameWithoutExtension($File)
$PDF      = $Path + '\' + $Filename  + '.pdf'

#Open File
$Word = New-Object -ComObject Word.Application
$Doc  = $Word.Documents.Open($File)

#Check Version of Word Installed and create PDF
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