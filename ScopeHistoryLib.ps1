<#
.Synopsis
�X�R�[�v �I�u�W�F�N�g�Ƃ��̋@�\�𗘗p���� Function ���C�u�����B
.Description
�X�R�[�v �I�u�W�F�N�g�Ƃ��̋@�\�𗘗p���� Function ���C�u�����B
�X�R�[�v �I�u�W�F�N�g�͔���J�I�u�W�F�N�g�ł���B
�]���āA�{���C�u�����̋@�\�͏����̃o�[�W�����œ��삵�Ȃ������قȂ铮������鋰�ꂪ����B
���s����ƁA���̂R�� Function ����`�����B
function Get-ScopeHistory
�X�R�[�v �I�u�W�F�N�g���擾����B
function Get-StrictMode
StrictMode �o�[�W�������擾����B�����ŃX�R�[�v �I�u�W�F�N�g���g�p����B
function IsGlobalScope
���s���̃X�R�[�v���O���[�o�� �X�R�[�v���ǂ������擾����B�����ŃX�R�[�v �I�u�W�F�N�g���g�p����B
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
    ���s���̃X�R�[�v�̗������擾����B
    .Description
    ���s���̃X�R�[�v��A���̐e�X�R�[�v�A���̂܂��e�̃X�R�[�v�A�c �Ƒk��A�O���[�o�� �X�R�[�v�܂ł��擾����B
    ����J�̋@�\���g�p���Ă��邽�߁A�����̃o�[�W�����œ��삵�Ȃ����ꂪ����B
    .Outputs
    �X�R�[�v �I�u�W�F�N�g�̃��X�g�B
    �擪�v�f�����s���̃X�R�[�v�B���X�g�̏��ɐe�X�R�[�v�ŁA�����̗v�f��
    �O���[�o�� �X�R�[�v�B
    .Example
    . Get-ScopeHistory
    ���s���̃X�R�[�v�̗������擾����B�B
    .Example 
    Get-ScopeHistory
    �ʏ�̌Ăяo���ł͎����I�Ɏq�X�R�[�v���쐬����邽�߁A���X�g�̐擪
    �͌Ăяo�����X�R�[�v�ł͂Ȃ��A���̎q�X�R�[�v�ɂȂ�B
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
    
    # �Ăяo�����Ɠ����X�R�[�v�ŌĂяo����Ă����O��Ԃ����L���Ȃ��悤�ɁA�q�X�R�[�v�𖾎�����B
    & {
        # SessionStateScope �� PowerShell �̌^�L���b�V���ŉ������Ȃ��B
        # Reflection ���g�p���� SessionStateScope �^�ƁA�W�F�l���b�N List �I�u�W�F�N�g �C���X�^���X���擾�B
        $scopetype = [System.AppDomain]::CurrentDomain.GetAssemblies() |% {$_.GetTypes() |? {$_.FullName -eq 'System.Management.Automation.SessionStateScope'}}
        $Result = ([System.Collections.Generic.List``1].MakeGenericType($scopetype)).GetConstructor([type]::EmptyTypes).Invoke(@())
        
        # �Ăяo�����X�R�[�v�擾�B
        # �q�X�R�[�v��؂����̂ŁA�P�e�X�R�[�v���Ăяo�����X�R�[�v�B
        $scope = $ExecutionContext `
            |% {$_.GetType().InvokeMember('_context',            [System.Reflection.BindingFlags] 'NonPublic, Instance, GetField', $null, $_, $null)} `
            |% {$_.GetType().InvokeMember('_engineSessionState', [System.Reflection.BindingFlags] 'NonPublic, Instance, GetField', $null, $_, $null)} `
            |% {$_.GetType().InvokeMember('currentScope',        [System.Reflection.BindingFlags] 'NonPublic, Instance, GetField', $null, $_, $null)} `
            |% {$_.GetType().InvokeMember( `
                    ($_.GetType().GetFields([System.Reflection.BindingFlags] 'NonPublic, Instance') |? {$_.Name -match 'Parent'}).Name, `
                    [System.Reflection.BindingFlags] 'NonPublic, Instance, GetField', $null, $_, $null
                    )
                }
        
        # ��c�X�R�[�v�̗�
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
    StrictMode �o�[�W�������擾����B
    .Description
    StrictMode �o�[�W�������擾����B
    ����J�̋@�\���g�p���Ă��邽�߁A�����̃o�[�W�����œ��삵�Ȃ����ꂪ����B
    .Parameter List
    ���ʂ��A�X�R�[�v�������Ƃ� StrictMode �o�[�W�����̃��X�g�Ƃ��ĕԂ��B
    �w�肵�Ȃ��ꍇ�A���s���̃X�R�[�v�ɓK�p����� StrictMode �o�[�W������Ԃ��B
    .Outputs
    StrictMode �o�[�W�����B
    -List �I�v�V�������w�肵���ꍇ�́AStrictMode �o�[�W�����̃��X�g�B
    ���X�g�̊e�v�f���X�R�[�v�ɉ����� StrictMode �o�[�W�������w���A�擪�v�f�����s���X�R�[�v�́A�����̗v�f���O���[�o�� �X�R�[�v�� StrictMode �o�[�W������\���B
    .Example
    Get-StrictMode
    ���s���̃X�R�[�v�� StrictMode �o�[�W�������擾����B
    .Example 
    . Get-StrictMode -List
    �X�R�[�v�̗������Ƃ� StrictMode �o�[�W�������擾����B
    .Example 
    Get-StrictMode -List
    �ʏ�̌Ăяo���ł͎����I�Ɏq�X�R�[�v���쐬����邽�߁A���X�g�̐擪
    �͌Ăяo�����X�R�[�v�ł͂Ȃ��A���̎q�X�R�[�v�ɂȂ�B
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
    �O���[�o�� �X�R�[�v�Ŏ��s���Ă��邱�Ƃ����؂���B
    .Description
    �O���[�o�� �X�R�[�v�Ŏ��s���Ă��邱�Ƃ����؂���B
    ����J�̋@�\���g�p���Ă��邽�߁A�����̃o�[�W�����œ��삵�Ȃ����ꂪ����B
    .Outputs
    �O���[�o�� �X�R�[�v�Ŏ��s���Ă���ꍇ�� True ��Ԃ��B����ȊO�̏ꍇ�� False ��Ԃ��B
    .Example
    . IsGlobalScope
    �O���[�o�� �X�R�[�v�Ŏ��s���Ă��邩���m�F����B
    .Example 
    . IsGlobalScope
    �O���[�o�� �X�R�[�v�Ŏ��s���̂̏ꍇ�� True ��Ԃ��B����ȊO�� False ��Ԃ��B
    .Example
    IsGlobalScope
    �ʏ�̌Ăяo���ł͎����I�Ɏq�X�R�[�v���쐬�����B�]���āAFunction ���O���[�o�� �X�R�[�v�Ŏ��s����邱�Ƃ͂Ȃ��A��� False ���Ԃ�B
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

    # �Ăяo�����Ɠ����X�R�[�v�ŌĂяo����Ă����O��Ԃ����L���Ȃ��悤�ɁA�q�X�R�[�v�𖾎�����B
    & {
        $Result = $false

        $scopehistory = (. Get-ScopeHistory)
        
        # �����I�Ɏq�X�R�[�v�ɓ����Ă���̂ŁA
        # �O���[�o�� �X�R�[�v���A2 �v�f�Ԃ�̂��z��B
        if($scopehistory.Count -eq 2) {
            $Result = $true
        } else {
            $Result = $false
        }
        Write-Output $Result
    }
}

New-Variable -Name _nasScopeHistoryLib_ -Value ([Version] '1.0.2.0') -Option ReadOnly -Force
