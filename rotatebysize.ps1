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
# 対象ファイルをサイズローテーション(切替えおよび削除)する。
function rotatebysize([uint32]$size, [uint32]$backup, [string]$filename, [string]$suffix) {

    # [実装メモ]
    # (1) 対象ファイルが所定のサイズに達した場合、過去ファイルへの切替えを実施する。
    # (2) 過去ファイルへの切替えを実施しなかった場合、以降の処理を実施しない。
    # (3) 対象ファイルを作成し直す。
    # (4) 保持世代数を超えた過去ファイルを削除する。

    $delim = "."
    $numpat = "^\d+$"

    # (1) 対象ファイルが所定のサイズに達した場合、過去ファイルへの切替えを実施する。
    $fname = "${filename}${suffix}"
    if (Test-Path -Path $fname) {
        if ($(Get-Item -Path $fname).Length -ge $size) {
            if (Test-Path -Path "${filename}${delim}$($backup)${suffix}") {
                Remove-Item -Path "${filename}${delim}$($backup)${suffix}" -ErrorAction Stop
            }
            for ($i = $backup - 1; $i -ge 1; $i -= 1) {
                if (Test-Path -Path "${filename}${delim}$($i)${suffix}") {
                    Rename-Item -Path "${filename}${delim}$($i)${suffix}" -NewName "${filename}${delim}$($i+1)${suffix}" -ErrorAction Stop
                }
            }
            Rename-Item -Path $fname -NewName "${filename}${delim}$(1)${suffix}" -ErrorAction Stop
        }
    }

    # (2) 過去ファイルへの切替えを実施しなかった場合、以降の処理を実施しない。
    if (Test-Path -Path $fname) {
        return
    }

    # (3) 対象ファイルを作成し直す。
    New-Item -ItemType File -Path $fname -ErrorAction Stop > $null

    # (4) 保持世代数を超えた過去ファイルを削除する。
    Get-Item -Path "${filename}${delim}*${suffix}" | ForEach-Object {
        $num = $_.Name.Substring($filename.Length + $delim.Length, $_.Name.Length - ($filename.Length + $delim.Length + $suffix.Length))
        if ($num -notmatch $numpat) {
            return
        }
        if ([uint32]$num -le $backup) {
            return
        }
        Remove-Item -Path $_ -ErrorAction Stop
    }
}

########################################################################
# サイズローテーションのテストコード。
$basedir = $(Split-Path -Path $MyInvocation.MyCommand.Path -Parent)
$logdir = $(New-Item -ItemType Directory -Path $(Join-Path -Path $basedir -ChildPath "log") -Force)
Push-Location -Path $logdir

for ($i = 0; $i -lt 100; $i += 1) {

    New-Item -ItemType File -Path "aaaa.log" -Value "aaaa $i" -Force > $null
    New-Item -ItemType File -Path "bbbb.log" -Value "bbbb $i" -Force > $null

    # サイズローテーション実行。
    rotatebysize 6 30 "aaaa.log"
    rotatebysize 6 30 "bbbb" ".log"

    # ローテーションされていないファイルを抽出。正しく動作すれば何も出力されない。
    Get-Item -Path "aaaa.log.*" | Where-Object {
        $i + 1 -ne [uint32]$(Get-Content -Path $_).Substring("aaaa ".Length) + [uint32]$_.Name.Substring("aaaa.log.".Length)
    }
    Get-Item -Path "bbbb.*.log" | Where-Object {
        $i + 1 -ne [uint32]$(Get-Content -Path $_).Substring("bbbb ".Length) + [uint32]$_.Name.Substring("bbbb.".Length, $_.Name.Length - ("bbbb.".Length + ".log".Length))
    }

    # 削除漏れを摘出。正しく動作すれば何も出力されない。
    Get-Item -Path "aaaa.log.*" | Where-Object {
        [uint32]$_.Name.Substring("aaaa.log.".Length) -gt 30
    }
    Get-Item -Path "bbbb.*.log" | Where-Object {
        [uint32]$_.Name.Substring("bbbb.".Length, $_.Name.Length - ("bbbb.".Length + ".log".Length)) -gt 30
    }
}

# 後始末。
Remove-Item -Path $(Get-Item -Path "aaaa.log*")
Remove-Item -Path $(Get-Item -Path "bbbb*.log")

Pop-Location
