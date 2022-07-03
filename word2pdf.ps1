$targetPath = ".\"

$files = Get-ChildItem -Path $targetPath | Where-Object { $_.Extension -like "*.docx" }

$word = New-Object -ComObject Word.Application
foreach ($f in $files) {
    Write-Host $f

    $doc = $word.Documents.Open($f.FullName)
    
    #$doc.ActiveWindow.View.ShowRevisionsAndComments = $False 

    $outputfile = $f.FullName.Replace("docx", "pdf")
    Write-Host $outputfile

    #https://docs.microsoft.com/ja-jp/office/vba/api/word.document.exportasfixedformat
    $doc.ExportAsFixedFormat($outputfile, 
        [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF,
        $False, #OpenAfterExport
        [Microsoft.Office.Interop.Word.WdExportOptimizeFor]::wdExportOptimizeForPrint,
        [Microsoft.Office.Interop.Word.WdExportRange]::wdExportAllDocument,
        0, #From
        0, #To
        [Microsoft.Office.Interop.Word.WdExportItem]::wdExportDocumentContent,
        $True, #IncludeDocProps
        $False, #KeepIRM
        [Microsoft.Office.Interop.Word.WdExportCreateBookmarks]::wdExportCreateHeadingBookmarks,
        $True, #DocStructureTags
        $True, #BitmapMissingFonts
        $False  #UseISO190051_
    )

    # https://docs.microsoft.com/ja-jp/office/vba/api/word.document.close(method)
    $doc.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)

}
# https://docs.microsoft.com/ja-jp/office/vba/api/word.application.quit(method)
$word.Quit()
Write-Host "done!"