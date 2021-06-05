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
# 対象ファイルを日次ローテーション(切替えおよび削除)する。
function rotatedaily([datetime]$dtm, [uint32]$backup, [string]$filename, [string]$suffix) {

    # [実装メモ]
    # (1) 対象ファイルの更新日付が現在日付と異なる場合、過去ファイルへの切替えを実施する。
    # (2) 過去ファイルへの切替えを実施しなかった場合、以降の処理を実施しない。
    # (3) 対象ファイルを作成し直す。
    # (4) 保持期間を超えた過去ファイルを削除する。

    $delim = "_"
    $dtfmt = "yyyyMMdd"
    $dtpat = "\d{4}\d{2}\d{2}"

    # (1) 対象ファイルの更新日付が現在日付と異なる場合、過去ファイルへの切替えを実施する。
    $fname = "${filename}${suffix}"
    if (Test-Path -Path $fname) {
        $ftime = $(Get-Item -Path $fname).LastWriteTime
        if ($ftime.Date -ne $dtm.Date) {
            $newname = "${filename}${delim}$($ftime.ToString($dtfmt))${suffix}"
            if (Test-Path -Path $newname) {
                Remove-Item -Path $newname -ErrorAction Stop
            }
            Rename-Item -Path $fname -NewName $newname -ErrorAction Stop
        }
    }

    # (2) 過去ファイルへの切替えを実施しなかった場合、以降の処理を実施しない。
    if (Test-Path -Path $fname) {
        return
    }

    # (3) 対象ファイルを作成し直す。
    New-Item -ItemType File -Path $fname -ErrorAction Stop > $null

    # (4) 保持期間を超えた過去ファイルを削除する。
    $threshold = $dtm.AddDays(- $backup).ToString($dtfmt)
    Get-Item -Path "${filename}${delim}*${suffix}" | ForEach-Object {
        if ($_.Name.Length -ne $filename.Length + $delim.Length + $dtfmt.Length + $suffix.Length) {
            return
        }
        $dt = $_.Name.Substring($filename.Length + $delim.Length, $dtfmt.Length)
        if ($dt -notmatch $dtpat) {
            return
        }
        if ($dt -ge $threshold) {
            return
        }
        Remove-Item -Path $_ -ErrorAction Stop
    }
}

########################################################################
# 日次ローテーションのテストコード。
$basedir = $(Split-Path -Path $MyInvocation.MyCommand.Path -Parent)
$logdir = $(New-Item -ItemType Directory -Path $(Join-Path -Path $basedir -ChildPath "log") -Force)
Push-Location -Path $logdir

# 前準備。
New-Item -ItemType File -Path "aaaa.log" -Force > $null
New-Item -ItemType File -Path "bbbb.log" -Force > $null
for ($i = -365; $i -le 0; $i += 1) {

    # 対象ファイル(前日分)。
    $now = $(Get-Date).AddDays($i)
    Set-ItemProperty -Path "aaaa.log" -Name LastWriteTime -Value $now.AddDays(-1)
    Set-ItemProperty -Path "bbbb.log" -Name LastWriteTime -Value $now.AddDays(-1)

    # 日次ローテーション実行。
    rotatedaily $now 30 "aaaa.log"
    rotatedaily $now 30 "bbbb" ".log"

    # 削除漏れを摘出。正しく動作すれば何も出力されない。
    $threshold = $now.AddDays(-30).ToString("yyyyMMdd")
    Get-Item -Path "aaaa.log_*" | Where-Object {
        $_.Name.Substring("aaaa.log_".Length, 8) -lt $threshold
    }
    Get-Item -Path "bbbb_*.log" | Where-Object {
        $_.Name.Substring("bbbb_".Length, 8) -lt $threshold
    }
}

# 後始末。
Remove-Item -Path $(Get-Item -Path "aaaa.log*")
Remove-Item -Path $(Get-Item -Path "bbbb*.log")

Pop-Location
