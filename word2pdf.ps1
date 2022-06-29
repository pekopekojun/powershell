$targetPath = ".\"

$files = Get-ChildItem -Path $targetPath | Where-Object { $_.Extension -like "*.docx" }

foreach ($f in $files) {
    Write-Host $f

    #Word�I�u�W�F�N�g�𐶐�
    $word = New-Object -ComObject Word.Application

    $doc = $word.Documents.Open($f.FullName)
    
    #�ύX�������\��
    #$doc.ActiveWindow.View.ShowRevisionsAndComments = $False 

    #�ۑ��t�@�C�����i�g���q��ύX�j
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

    $doc.Close()
    $word.Quit()
}
Write-Host "done!"