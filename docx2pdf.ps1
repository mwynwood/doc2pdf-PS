
# Bulk converts Word Documents to PDFs
# Thanks to: https://stackoverflow.com/a/46287197 

$documents_path=$args[0]
# $documents_path = 'C:\documents'
$word_app = New-Object -ComObject Word.Application
Get-ChildItem -Path $documents_path -Filter *.doc? | ForEach-Object {
    $document = $word_app.Documents.Open($_.FullName)
    Write-Host -NoNewline "Converting:" $_.FullName
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    $document.SaveAs([ref] $pdf_filename, [ref] 17)
    $document.Close()
    Write-Host " Done."
}
$word_app.Quit()