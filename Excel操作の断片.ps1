#
# Copyright 2021 agwlvssainokuni
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#

########################################################################
# PowerShellスクリプト小品集
########################################################################

########################################################################
# Excel操作の断片

# Excelを開く。
$excel = New-Object -ComObject Excel.Application -Property @{Visible=$true;DisplayAlerts=$false}

# 引数指定を省略する場合は以下を受け渡す。
[System.Reflection.Missing]::Value

# 引数に列挙型の値を指定する場合は以下の記法を使う。
[Microsoft.Office.Interop.Excel.XlPlatform]::xlMacintosh
[Microsoft.Office.Interop.Excel.XlPlatform]::xlMSDOS
[Microsoft.Office.Interop.Excel.XlPlatform]::xlWindows

[Microsoft.Office.Interop.Excel.XlDirection]::xlDown
[Microsoft.Office.Interop.Excel.XlDirection]::xlUp
[Microsoft.Office.Interop.Excel.XlDirection]::xlToLeft
[Microsoft.Office.Interop.Excel.XlDirection]::xlToRight

[Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteValues

# 既存のブックを開く。
# https://docs.microsoft.com/ja-jp/office/vba/api/excel.workbooks.open
$book = $excel.Workbooks.Open($filepath)

# 新規ブックを作成する。
# https://docs.microsoft.com/ja-jp/office/vba/api/excel.workbooks.add
$book = $excel.Workbooks.Add()

# 既存シートを選択する。
# コレクションの添字は「1」始まり。
$sheet = $book.Worksheets.Item(1)
# シート名でも指定できる。
$sheet = $book.Worksheets.Item($sheetname)

# 新規シートを追加する。
# https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheets.add
# 現在のシートの直前に追加。
$sheet = $book.Worksheets.Add()
# 先頭に追加。
$sheet = $book.Worksheets.Add($book.Worksheets.Item(1))
# 末尾に追加。
$sheet = $book.Worksheets.Add([System.Reflection.Missing]::Value, $book.Worksheets.Item($book.Worksheets.Count))

# シートを表示する。
$sheet.Select()

# シートを削除する。
# https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.delete
$sheet.Delete()

# ブックを保存する。
# https://docs.microsoft.com/ja-jp/office/vba/api/excel.workbook.save
# https://docs.microsoft.com/ja-jp/office/vba/api/excel.workbook.saveas
$book.Save()
$book.SaveAs($filepath)

# ブックを閉じる。
# https://docs.microsoft.com/ja-jp/office/vba/api/excel.workbook.close
$book.Close()

# Excelを終了する。
$excel.Quit()
$excel = $null
[GC]::Collect()
