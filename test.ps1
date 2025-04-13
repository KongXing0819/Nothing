param(
    [int]$frequency = 11000,
    [int]$groupNumber = 3,
    [switch]$Random
)

# 版本兼容层 ============================================
if (-not (Test-Path variable:IsWindows)) {
    $IsWindows = $env:OS -eq 'Windows_NT'
    try {
        $IsWindows = [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
            [System.Runtime.InteropServices.OSPlatform]::Windows
        )
    } catch {}
}

# 数据定义 ==============================================
$groupMembers = @{
    '1'  = @("谢智稳", "匡文享", "邓梦琪", "肖世杰")
    '2'  = @("刘梓暄", "王鸿毅", "邓雅丽", "李嘉倪")
    '3'  = @("张涵娇", "李米", "邓高誉", "申谨萱")
    '4'  = @("赵鸿聪", "李金洛", "吕方熙", "李勇慧")
    '5'  = @("尹香婧", "叶俊宇", "李熙峰", "黄文俊")
    '6'  = @("李文睿", "李文磊", "李浩然", "曾紫轩")
    '7'  = @("周丹", "金子超", "刘可欣", "朱骐可")
    '8'  = @("龙鸿运", "朱靓颖", "罗思颖", "李德祥")
    '9'  = @("李佳蕊", "刘福源", "李江林", "潘棋圆")
    '10' = @("周硕彦", "刘佳欣", "刘佳淼", "禹浩")
    '11' = @("宁振义", "李姝祺", "尹锦涵", "黄瑞")
}

try {
    # 输入验证 ==========================================
    if ($groupNumber -lt 1) { throw "组数不能小于1" }
    $validGroups = $groupMembers.Keys | Sort-Object { [int]$_ }
    $totalGroups = $validGroups.Count
    if ($groupNumber -gt $totalGroups) {
        throw "需要选择的组数 ($groupNumber) 超过存在的组数 ($totalGroups)"
    }

    # 清理目录 ==========================================
    $mathFolder = Join-Path -Path $PSScriptRoot -ChildPath "数学"
    if (Test-Path $mathFolder) {
        Remove-Item -Path $mathFolder -Recurse -Force -ErrorAction Stop
        Start-Sleep -Milliseconds 500
    }

    # 随机数生成 ========================================
    if ($Random) {
        $rnd = [System.Random]::new()
        Write-Host "`n这次使用完全随机生成`n"
    } else {
        $dateStr = Get-Date -Format "yyyyMMdd"
        $seed = [BitConverter]::ToInt32(
            ([System.Security.Cryptography.SHA1]::Create().ComputeHash(
                [Text.Encoding]::UTF8.GetBytes($dateStr)
            ))[0..3], 0
        )
        $rnd = [System.Random]::new($seed)
        Write-Host "`n用于随机种子的日期: $(Get-Date -Format 'yyyy年MM月dd日')`n"
    }

    # 统计分析 ==========================================
    $stats = 1..$frequency | ForEach-Object { 
        $validGroups[$rnd.Next(0, $totalGroups)] 
    } | Group-Object | ForEach-Object {
        [PSCustomObject]@{
            Number = [int]$_.Name
            Count = $_.Count
        }
    } | Sort-Object @{e='Count'; Descending=$true}, @{e='Number'; Ascending=$true}

    $stats = $stats | ForEach-Object {
        [PSCustomObject]@{
            组号     = $_.Number
            频数    = $_.Count
            占比百分比 = [math]::Round(($_.Count / $frequency) * 100, 2)
        }
    }

    Write-Host "随机数统计 (样本数: $frequency)" -ForegroundColor Cyan
    $stats | Format-Table -AutoSize -Property @(
        @{l="组"; e={$_.组号}}
        @{l="频数(次)"; e={$_.频数}; Align='Right'}
        @{l="占比(%)"; e={$_.占比百分比}; Align='Right'}
    )

    # 结果输出 ==========================================
    $topGroups = $stats | Select-Object -First $groupNumber
    Write-Host "`n要交作业的组为:" -ForegroundColor Green
    $topGroups | ForEach-Object { Write-Host "第$($_.组号)组" }

    $allMembers = $topGroups.组号 | Sort-Object | ForEach-Object { 
        $groupMembers[$_.ToString()] 
    } | Sort-Object -Culture "zh-CN" -Unique

    Write-Host "`n要交作业的人员:" -ForegroundColor Green
    Write-Host ($allMembers -join '、')

    # 文件生成 ==========================================
    New-Item -Path $mathFolder -ItemType Directory -Force | Out-Null
    
    $encoding = if ($IsWindows) { 
        [System.Text.Encoding]::GetEncoding(
            [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage
        )
    } else { 
        [System.Text.Encoding]::UTF8 
    }

    $allMembers | ForEach-Object {
        $fileName = if ($IsWindows) { "$_.bat" } else { "$_.sh" }
        $filePath = Join-Path -Path $mathFolder -ChildPath $fileName
        
        if ($IsWindows) {
@"
@echo off
chcp 65001 > nul
( ping 127.0.0.1 -n 2 > nul && del /f "%0" ) & exit
"@ | Out-File -FilePath $filePath -Encoding $encoding
        } else {
@'
#!/bin/bash
sleep 0.5
rm -f "$0"
exit 0
'@ | Out-File -FilePath $filePath -Encoding utf8
            chmod +x $filePath
        }
    }
    Write-Host "`n生成文件编码格式: $($encoding.EncodingName)" -ForegroundColor Cyan
}
catch {
    Write-Host "发生错误: $_" -ForegroundColor Red
    exit 1
}
finally {
    # 跨平台退出处理 ====================================
    $isInteractive = [Environment]::UserInteractive -and 
                    (-not [Console]::IsInputRedirected) -and 
                    (-not [Console]::IsOutputRedirected)

    if ($isInteractive) {
        if ($IsWindows) {
            Write-Host "`n按任意键退出..." -NoNewline
            try {
                $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            } catch {
                Start-Sleep 3
            }
        } else {
            Write-Host
        }
    }
}
