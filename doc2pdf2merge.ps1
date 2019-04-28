# doc2pdf2merge
# This script converts all the Word Documents in a folder to PDF files.
# Next, it creates a simple cover page PDF.
# It then merges all the PDFs into one PDF file.
#
# The order is dertermined by alphabetical order, so if you name your 
# Word Docuemnts with a number at the start, it'll get them in your
# desired order.
# 
# Created by Marcus Wynwood
# 28 April 2019
# https://github.com/mwynwood/doc2pdf 
#
# Thanks to:
#  https://stackoverflow.com/a/46287197
#  http://www.pdfsharp.net/PDFsharp_License.ashx
#  https://github.com/mikepfeiffer/PowerShell/blob/master/Merge-PDF.ps1

Add-Type -Path .\PdfSharp.dll
Add-Type -AssemblyName System.Windows.Forms

# This function does all the tings.
# Pass it a folder full of Word Docs, some text for the Cover Page, and an Image.
Function doc2pdf2merge ($documents_path, $coverPageText1, $coverPageText2, $coverPageText3, $coverPageText4, $coverPageImage) {
        
    $nameOfMergedPDF = "merged.pdf"
    $nameOfCoverPage = "!!!CoverPage.pdf"

    $counter = 0
    $documents_path = $FolderBrowser.SelectedPath

    # Convert DOCs to PDFs
    $word_app = New-Object -ComObject Word.Application
    Get-ChildItem -Path $documents_path -Filter *.doc? | ForEach-Object {
        $document = $word_app.Documents.Open($_.FullName)
        Write-Host "Converting:" $_.FullName
        $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
        $document.SaveAs([ref] $pdf_filename, [ref] 17)
        $document.Close()
        $counter++
    }
    $word_app.Quit()
    Write-Host "$counter Word Documents have been converted to PDF. `r`n"

    # Create Cover Page to prepend
    Write-Host "Generating Cover Page..."
    $doc = New-Object PdfSharp.Pdf.PdfDocument
    $doc.Info.Title = "Cover Page"
    $page = $doc.AddPage()
    $gfx = [PdfSharp.Drawing.XGraphics]::FromPdfPage($page)
    $font = New-Object PdfSharp.Drawing.XFont("Calibri", 20, [PdfSharp.Drawing.XFontStyle]::Bold)
    
    $rect = New-Object PdfSharp.Drawing.XRect(0,0,$page.Width, $page.Height)
    $gfx.DrawString($coverPageText1, $font, [PdfSharp.Drawing.XBrushes]::Black, $rect, [PdfSharp.Drawing.XStringFormats]::Center)

    $rect = New-Object PdfSharp.Drawing.XRect(0,40,$page.Width, $page.Height)
    $gfx.DrawString($coverPageText2, $font, [PdfSharp.Drawing.XBrushes]::Black, $rect, [PdfSharp.Drawing.XStringFormats]::Center)

    $rect = New-Object PdfSharp.Drawing.XRect(0,80,$page.Width, $page.Height)
    $gfx.DrawString($coverPageText3, $font, [PdfSharp.Drawing.XBrushes]::Black, $rect, [PdfSharp.Drawing.XStringFormats]::Center)

    $rect = New-Object PdfSharp.Drawing.XRect(0,120,$page.Width, $page.Height)
    $gfx.DrawString($coverPageText4, $font, [PdfSharp.Drawing.XBrushes]::Black, $rect, [PdfSharp.Drawing.XStringFormats]::Center)

    $image = [PdfSharp.Drawing.XImage]::FromFile($coverPageImage)
    $gfx.DrawImage($image, 200, 200, 200,200)
    
    $doc.Save($documents_path + "\" + $nameOfCoverPage)

    # Do the Merge of all the PDFs
    Write-Host "Merging all the PDF files into one PDF file..."
    $output = New-Object PdfSharp.Pdf.PdfDocument
    $PdfReader = [PdfSharp.Pdf.IO.PdfReader]
    $PdfDocumentOpenMode = [PdfSharp.Pdf.IO.PdfDocumentOpenMode]
    foreach($i in (Get-ChildItem $documents_path *.pdf)) {
        $input = New-Object PdfSharp.Pdf.PdfDocument
        $input = $PdfReader::Open($i.fullname, $PdfDocumentOpenMode::Import)
        $input.Pages | %{$output.AddPage($_)}
    }
    $output.Save($documents_path + "\" + $nameOfMergedPDF)
    Start-Process $documents_path
}

# Let's Do It
$title1 = Read-Host -Prompt "Enter Cover Page title 1"
$title2 = Read-Host -Prompt "Enter Cover Page title 2"
$title3 = Read-Host -Prompt "Enter Cover Page title 3"
$title4 = Read-Host -Prompt "Enter Cover Page title 4"
#TODO: Make this an array
$img = $PSScriptRoot + "\logo.png"
Write-Host "Select the folder that contains the Word Documents..."

$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$FolderBrowser.ShowNewFolderButton = $false;
$FolderBrowser.Description = "Select the folder that contains the Word Documents"

if($FolderBrowser.ShowDialog() -eq "OK") {
    doc2pdf2merge $FolderBrowser.SelectedPath $title1 $title2 $title3 $title4 $img
    $FolderBrowser.Dispose()
}
