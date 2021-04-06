<#
.Synopsis
スコープ オブジェクトとその機能を利用する Function ライブラリ。
.Description
スコープ オブジェクトとその機能を利用する Function ライブラリ。
スコープ オブジェクトは非公開オブジェクトである。
従って、本ライブラリの機能は将来のバージョンで動作しない恐れや異なる動作をする恐れがある。
実行すると、次の３つの Function が定義される。
function Get-ScopeHistory
スコープ オブジェクトを取得する。
function Get-StrictMode
StrictMode バージョンを取得する。内部でスコープ オブジェクトを使用する。
function IsGlobalScope
実行中のスコープがグローバル スコープかどうかを取得する。内部でスコープ オブジェクトを使用する。
.Link
https://www.nassystems.info/downloads/Microsoft/Windows/PowerShell/ScopeHistoryLib.ps1
.Notes
ScopeHistoryLib version 1.02
    
MIT License
    
Copyright (c) 2018 NASsystems.info
    
Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:
    
The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.
    
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

function Get-ScopeHistory {
    <#
    .Synopsis
    実行中のスコープの履歴を取得する。
    .Description
    実行中のスコープや、その親スコープ、そのまた親のスコープ、… と遡り、グローバル スコープまでを取得する。
    非公開の機能を使用しているため、将来のバージョンで動作しない恐れがある。
    .Outputs
    スコープ オブジェクトのリスト。
    先頭要素が実行中のスコープ。リストの順に親スコープで、末尾の要素が
    グローバル スコープ。
    .Example
    . Get-ScopeHistory
    実行中のスコープの履歴を取得する。。
    .Example 
    Get-ScopeHistory
    通常の呼び出しでは自動的に子スコープが作成されるため、リストの先頭
    は呼び出し元スコープではなく、その子スコープになる。
    .LINK
    https://www.nassystems.info/downloads/Microsoft/Windows/PowerShell/ScopeHistoryLib.ps1
    .Notes
    ScopeHistoryLib version 1.02
    
    MIT License
    
    Copyright (c) 2018 NASsystems.info
    
    Permission is hereby granted, free of charge, to any person obtaining a
    copy of this software and associated documentation files (the
    "Software"), to deal in the Software without restriction, including
    without limitation the rights to use, copy, modify, merge, publish,
    distribute, sublicense, and/or sell copies of the Software, and to
    permit persons to whom the Software is furnished to do so, subject to
    the following conditions:
    
    The above copyright notice and this permission notice shall be included
    in all copies or substantial portions of the Software.
    
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
    OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
    MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
    IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
    CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
    TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
    SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #>
    
    # 呼び出し元と同じスコープで呼び出されても名前空間を共有しないように、子スコープを明示する。
    & {
        # SessionStateScope は PowerShell の型キャッシュで解決しない。
        # Reflection を使用して SessionStateScope 型と、ジェネリック List オブジェクト インスタンスを取得。
        $scopetype = [System.AppDomain]::CurrentDomain.GetAssemblies() |% {$_.GetTypes() |? {$_.FullName -eq 'System.Management.Automation.SessionStateScope'}}
        $Result = ([System.Collections.Generic.List``1].MakeGenericType($scopetype)).GetConstructor([type]::EmptyTypes).Invoke(@())
        
        # 呼び出し元スコープ取得。
        # 子スコープを切ったので、１つ親スコープが呼び出し元スコープ。
        $scope = $ExecutionContext `
            |% {$_.GetType().InvokeMember('_context',            [System.Reflection.BindingFlags] 'NonPublic, Instance, GetField', $null, $_, $null)} `
            |% {$_.GetType().InvokeMember('_engineSessionState', [System.Reflection.BindingFlags] 'NonPublic, Instance, GetField', $null, $_, $null)} `
            |% {$_.GetType().InvokeMember('currentScope',        [System.Reflection.BindingFlags] 'NonPublic, Instance, GetField', $null, $_, $null)} `
            |% {$_.GetType().InvokeMember( `
                    ($_.GetType().GetFields([System.Reflection.BindingFlags] 'NonPublic, Instance') |? {$_.Name -match 'Parent'}).Name, `
                    [System.Reflection.BindingFlags] 'NonPublic, Instance, GetField', $null, $_, $null
                    )
                }
        
        # 先祖スコープの列挙
        while($scope) {
            $Result.Add($scope) | Out-Null
            $scope = $scope.GetType().InvokeMember( `
                ($scope.GetType().GetFields([System.Reflection.BindingFlags] 'NonPublic, Instance') |? {$_.Name -match 'Parent'}).Name, `
                [System.Reflection.BindingFlags] 'NonPublic, Instance, GetField', $null, $scope, $null
                )
        }
        
        Write-Output (,$Result.AsReadOnly())
    }
}


function Get-StrictMode {
    <#
    .Synopsis
    StrictMode バージョンを取得する。
    .Description
    StrictMode バージョンを取得する。
    非公開の機能を使用しているため、将来のバージョンで動作しない恐れがある。
    .Parameter List
    結果を、スコープ履歴ごとの StrictMode バージョンのリストとして返す。
    指定しない場合、実行中のスコープに適用される StrictMode バージョンを返す。
    .Outputs
    StrictMode バージョン。
    -List オプションを指定した場合は、StrictMode バージョンのリスト。
    リストの各要素がスコープに於ける StrictMode バージョンを指し、先頭要素が実行中スコープの、末尾の要素がグローバル スコープの StrictMode バージョンを表す。
    .Example
    Get-StrictMode
    実行中のスコープの StrictMode バージョンを取得する。
    .Example 
    . Get-StrictMode -List
    スコープの履歴ごとの StrictMode バージョンを取得する。
    .Example 
    Get-StrictMode -List
    通常の呼び出しでは自動的に子スコープが作成されるため、リストの先頭
    は呼び出し元スコープではなく、その子スコープになる。
    .LINK
    https://www.nassystems.info/downloads/Microsoft/Windows/PowerShell/ScopeHistoryLib.ps1
    .Notes
    ScopeHistoryLib version 1.02
    
    MIT License
    
    Copyright (c) 2018 NASsystems.info
    
    Permission is hereby granted, free of charge, to any person obtaining a
    copy of this software and associated documentation files (the
    "Software"), to deal in the Software without restriction, including
    without limitation the rights to use, copy, modify, merge, publish,
    distribute, sublicense, and/or sell copies of the Software, and to
    permit persons to whom the Software is furnished to do so, subject to
    the following conditions:
    
    The above copyright notice and this permission notice shall be included
    in all copies or substantial portions of the Software.
    
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
    OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
    MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
    IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
    CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
    TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
    SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #>
    param([switch] $List)
    $scopehistory = . Get-ScopeHistory

    if($List.IsPresent) {
        $Result = New-Object System.Collections.Generic.List[System.Version]
        foreach($scope in $scopehistory.GetEnumerator()) {
            $strictmodeversion = $scope.GetType().InvokeMember( `
                ($scope.GetType().GetFields([System.Reflection.BindingFlags] 'NonPublic, Instance') |? {$_.Name -match 'StrictModeVersion'}).Name, `
                [System.Reflection.BindingFlags] 'NonPublic, Instance, GetField', $null, $scope, $null
                )
            $Result.Add($strictmodeversion)
        }
        $Result = $Result.AsReadOnly()
    } else {
        foreach($scope in $scopehistory.GetEnumerator()){
            $strictmodeversion = $scope.GetType().InvokeMember( `
                ($scope.GetType().GetFields([System.Reflection.BindingFlags] 'NonPublic, Instance') |? {$_.Name -match 'StrictModeVersion'}).Name, `
                [System.Reflection.BindingFlags] 'NonPublic, Instance, GetField', $null, $scope, $null
                )
            if($strictmodeversion) {
                $Result = $strictmodeversion
                break
            }
        }
    }
    
    ,$Result
}


function IsGlobalScope {
    <#
    .Synopsis
    グローバル スコープで実行していることを検証する。
    .Description
    グローバル スコープで実行していることを検証する。
    非公開の機能を使用しているため、将来のバージョンで動作しない恐れがある。
    .Outputs
    グローバル スコープで実行している場合は True を返す。それ以外の場合は False を返す。
    .Example
    . IsGlobalScope
    グローバル スコープで実行しているかを確認する。
    .Example 
    . IsGlobalScope
    グローバル スコープで実行中のの場合は True を返す。それ以外は False を返す。
    .Example
    IsGlobalScope
    通常の呼び出しでは自動的に子スコープが作成される。従って、Function がグローバル スコープで実行されることはなく、常に False が返る。
    .Notes
    ScopeHistoryLib version 1.02
    
    MIT License
    
    Copyright (c) 2018 NASsystems.info
    
    Permission is hereby granted, free of charge, to any person obtaining a
    copy of this software and associated documentation files (the
    "Software"), to deal in the Software without restriction, including
    without limitation the rights to use, copy, modify, merge, publish,
    distribute, sublicense, and/or sell copies of the Software, and to
    permit persons to whom the Software is furnished to do so, subject to
    the following conditions:
    
    The above copyright notice and this permission notice shall be included
    in all copies or substantial portions of the Software.
    
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
    OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
    MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
    IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
    CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
    TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
    SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #>

    # 呼び出し元と同じスコープで呼び出されても名前空間を共有しないように、子スコープを明示する。
    & {
        $Result = $false

        $scopehistory = (. Get-ScopeHistory)
        
        # 明示的に子スコープに入っているので、
        # グローバル スコープ時、2 要素返るのが想定。
        if($scopehistory.Count -eq 2) {
            $Result = $true
        } else {
            $Result = $false
        }
        Write-Output $Result
    }
}

New-Variable -Name _nasScopeHistoryLib_ -Value ([Version] '1.0.2.0') -Option ReadOnly -Force
