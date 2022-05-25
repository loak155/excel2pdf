# 実行コマンド(xlsx2pdf.bat)
# powershell -NoProfile -ExecutionPolicy Unrestricted .\xlsx2pdf.ps1 %1

if ( $args -eq $null ) {
    Write-Error '引数がありません'
}

if (Test-Path $args[0]) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    # 拡張子変更
    $newpdf = $args[0] -replace '\.xlsx$', '.pdf'


    try {
        $book = $excel.Workbooks.Open($args[0])

        # PDF
        $book.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $newpdf)

        $book.Close($false)

    }
    catch {
        Write-Error 'エラーが発生しました'
    }
    finally {
        $excel.Quit()
        $excel = $null
        [GC]::Collect()
    }

}
else {
    Write-Error 'ファイルが見つかりません'
}
