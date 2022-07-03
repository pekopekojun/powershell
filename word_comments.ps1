$targetPath = ".\"

$files = Get-ChildItem -Path $targetPath | Where-Object { $_.Extension -like "*.docx" }

$word = New-Object -ComObject Word.Application
foreach ($f in $files) {
    Write-Host $f

    $doc = $word.Documents.Open($f.FullName)
    
    #$doc.ActiveWindow.View.ShowRevisionsAndComments = $False 

    $outputfile = $f.FullName.Replace("docx", "pdf")
    Write-Host $outputfile
    # https://docs.microsoft.com/ja-jp/office/vba/api/word.comment
    foreach ($c in $doc.Comments) {
        Write-Host $c.Date 
        Write-Host $c.Author 
        Write-Host $c.Scope.Information([Microsoft.Office.Interop.Word.WdInformation]::wdActiveEndPageNumber)
        Write-Host $c.Scope.Information([Microsoft.Office.Interop.Word.WdInformation]::wdFirstCharacterLineNumber)
        Write-Host $c.Scope.Text
        Write-Host $c.Range.Text 
    }

    # https://docs.microsoft.com/ja-jp/office/vba/api/word.document.close(method)
    $doc.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)

}
# https://docs.microsoft.com/ja-jp/office/vba/api/word.application.quit(method)
$word.Quit()
Write-Host "done!"