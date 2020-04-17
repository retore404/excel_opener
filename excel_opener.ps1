######################################################################
## スクリプト名：excel_opener.ps1
## スクリプト内容：引数で渡すcsvに記載されている文字列で
## 　　　　　　　　パスワード付きExcelを開けないか試すスクリプト．
## 　　　　　　　　試行の結果は，同フォルダにログとして出力する．
## 引数：対象のExcelファイルパス, パスワード候補のcsvファイルパス
## 戻り値：なし（同フォルダにログを出力）
######################################################################

# 引数の受け取り
$filepath = $Args[0] # ターゲットExcelファイルパス
$password_list_path = $Args[1] # パスワード候補csvファイルパス

# csvからパスワードリストを読み取り
$password_list = Import-Csv $password_list_path -Encoding UTF8
$password_list | ConvertFrom-Csv -Header @('password')
$password_list | Format-Table

foreach($p in $password_list){
    try { 
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $true
    
        # パスワードを指定してブックを開く
        $wb = $excel.Workbooks.Open($filepath, [Type]::Missing, [Type]::Missing, [Type]::Missing, $p.password)
        $excel.Quit()

        $p.password
    
    } finally {
        $sheet, $wb, $excel | ForEach-Object {
            if ($_ -ne $null) {
                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
            }
        }
    }
}




