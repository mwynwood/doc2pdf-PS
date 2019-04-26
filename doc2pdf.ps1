# doc2pdf-gui
# A script to bulk convert DOC and DOCX files to PDF
# 
# Created by Marcus Wynwood
# 26 April 2019
# https://github.com/mwynwood/doc2pdf 
#
# Thanks to: https://stackoverflow.com/a/46287197

Add-Type -AssemblyName System.Windows.Forms

$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$FolderBrowser.ShowNewFolderButton = $false;
$FolderBrowser.Description = "Select the folder that contains the Word Documents"

$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = '800,400'
$Form.text = "doc2pdf"
$Form.TopMost = $true

$TextBox1 = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline = $true
$TextBox1.WordWrap = $true
$TextBox1.width = 800
$TextBox1.height = 400
$TextBox1.location = New-Object System.Drawing.Point(0,0)
$TextBox1.Font = 'Microsoft Sans Serif,10'
$TextBox1.Scrollbars = "Vertical"

$Form.controls.AddRange(@($TextBox1))

if($FolderBrowser.ShowDialog() -eq "OK") {
    $counter = 0
    $Form.Show()
    $documents_path = $FolderBrowser.SelectedPath
    $word_app = New-Object -ComObject Word.Application
    Get-ChildItem -Path $documents_path -Filter *.doc? | ForEach-Object {
        $document = $word_app.Documents.Open($_.FullName)
        $TextBox1.AppendText("Converting: " + $_.FullName + "`r`n")
        $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
        $document.SaveAs([ref] $pdf_filename, [ref] 17)
        $document.Close()
        $TextBox1.AppendText("Done. `r`n `r`n")
        $counter++
    }
    $word_app.Quit()
    $TextBox1.AppendText("$counter Word Documents have been converted to PDF. `r`n `r`n")
    [void][System.Windows.Forms.MessageBox]::Show("$counter Word Documents have been converted to PDF.`r`nClick OK to open the folder.", "doc2pdf Conversion Complete", [System.Windows.Forms.MessageBoxButtons]::OK)
    # $TextBox1.AppendText("Opening folder in ")
    # For ($i=5; $i -ge 0; $i--) {
    #     $TextBox1.AppendText($i.ToString() + ".. ")
    #     Start-Sleep -Seconds 1
    # }
    $Form.Hide()
    Start-Process $documents_path
} else {
    Write-Host "No folder selected. Quitting."
}
$FolderBrowser.Dispose()