######################################################################
## スクリプト名：excel_opener.ps1
## スクリプト内容：引数で渡すcsvに記載されている文字列で
## 　　　　　　　　パスワード付きExcelを開けないか試すスクリプト．
## 　　　　　　　　試行の結果は，標準出力する．
## 引数：対象のExcelファイルパス, パスワード候補のcsvファイルパス
## 戻り値：なし
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
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
    
        # パスワードを指定してブックを開く
        $wb = $null # wbを初期化
        $wb = $excel.Workbooks.Open($filepath, [Type]::Missing, [Type]::Missing, [Type]::Missing, $p.password)
        $excel.Quit()

        # 正しいパスワードを標準出力
        Write-Host "The correct password is " $p.password
        break
    } catch {
        Write-Host $p.password "is a wrong password."
        $excel.quit()
    } finally {
        $sheet, $wb, $excel | ForEach-Object {
            if ($_ -ne $null) {
                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
            }
        }
    }
}




