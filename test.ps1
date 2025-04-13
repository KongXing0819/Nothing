param(
    [int]$frequency = 11000,
    [int]$groupNumber = 3,
    [switch]$ForceRandom
)

$jsonPath = Join-Path $PSScriptRoot "groups.json"

# 读取并初始化JSON结构
try {
    $jsonContent = Get-Content $jsonPath -Encoding UTF8 -Raw | ConvertFrom-Json -NoEnumerate
    
    # 初始化必要字段
    if (-not $jsonContent.PSObject.Properties['Time']) {
        $jsonContent | Add-Member -MemberType NoteProperty -Name 'Time' -Value ""
    }

    foreach ($groupKey in ($jsonContent.PSObject.Properties.Name | Where-Object { $_ -ne 'Time' })) {
        $group = $jsonContent.$groupKey
        
        if (-not $group.PSObject.Properties['Selected']) {
            $group | Add-Member -MemberType NoteProperty -Name 'Selected' -Value $false
        }
        if (-not $group.PSObject.Properties['LastSelected']) {
            $group | Add-Member -MemberType NoteProperty -Name 'LastSelected' -Value $false
        }
        
        $group.weight = [double]$group.weight
    }
}
catch {
    throw "JSON文件处理失败: $_"
}

# 日期比较逻辑
$currentDate = Get-Date -Format "yyyyMMdd"
$sameDay = $jsonContent.Time -eq $currentDate
$selectedCount = ($jsonContent.PSObject.Properties | Where-Object { 
    $_.Name -ne 'Time' -and $_.Value.Selected 
}).Count

# 处理选中组不足的情况
if (-not $ForceRandom -and $sameDay) {
    if ($selectedCount -eq 0) {
        Write-Host "未检测到已选组，将执行随机选择..." -ForegroundColor Yellow
        $sameDay = $false
    }
    elseif ($selectedCount -lt $groupNumber) {
        Write-Host "检测到选中组数不足（$selectedCount/$groupNumber），重置所有组状态..." -ForegroundColor Yellow
        $jsonContent.PSObject.Properties | Where-Object Name -ne 'Time' | ForEach-Object {
            $_.Value.Selected = $false
            $_.Value.LastSelected = $false
        }
        $jsonContent.Time = ""
        $jsonContent | ConvertTo-Json -Depth 3 | Set-Content $jsonPath -Encoding UTF8
        $sameDay = $false
    }
    else {
        # 正常输出已选组
        $selectedGroups = @($jsonContent.PSObject.Properties | 
            Where-Object { 
                $_.Name -ne 'Time' -and 
                $_.Value.Selected 
            } |
            Sort-Object { [int]$_.Name } | 
            Select-Object -First $groupNumber)
        
        Write-Host "`n要交作业的组为:" -ForegroundColor Green
        $selectedGroups | ForEach-Object { 
            Write-Host "第$($_.Name)组"
        }
        
        # 输出成员名单
        $allMembers = @($selectedGroups | ForEach-Object { 
            $_.Value.Members 
        } | Sort-Object -Culture "zh-CN" -Unique)
        
        Write-Host "`n要交作业的人员:" -ForegroundColor Green
        Write-Host ($allMembers -join '、')
        
        exit 0
    }
}

# 执行加权随机逻辑（关键修改：严格排除已选组）
try {
    # 重置所有组状态（确保开始随机选择前状态正确）
    $jsonContent.PSObject.Properties | Where-Object Name -ne 'Time' | ForEach-Object {
        $_.Value.LastSelected = $_.Value.Selected
        $_.Value.Selected = $false
    }

    # 生成有效组列表（严格排除已选组）
    $validGroups = @($jsonContent.PSObject.Properties | 
        Where-Object { 
            $_.Name -ne 'Time' -and 
            $_.Value.weight -gt 0 -and
            -not $_.Value.LastSelected  # 新增：排除上次选中的组
        } | 
        Sort-Object { [int]$_.Name })

    # 权重计算
    $totalWeight = 0
    $groupList = @($validGroups | ForEach-Object {
        $group = $_.Value
        $totalWeight += $group.weight
        
        [PSCustomObject]@{
            Name    = $_.Name
            Number  = [int]$_.Name
            Weight  = $group.weight
            Members = $group.Members
            Object  = $group
        }
    } | Sort-Object Number)

    # 输入验证
    if ($totalWeight -le 0) { throw "没有有效的组可供选择（所有组权重≤0或已被选中）" }
    if ($groupNumber -gt $groupList.Count) {
        throw "需要选择的组数 ($groupNumber) 超过有效组数 ($($groupList.Count))"
    }

    # 随机数生成
    $rnd = if ($ForceRandom) { 
        [System.Random]::new() 
    } else {
        $seed = [BitConverter]::ToInt32(
            (([System.Security.Cryptography.SHA1]::Create()).ComputeHash(
                [Text.Encoding]::UTF8.GetBytes($currentDate)
            ))[0..3], 0
        )
        [System.Random]::new($seed)
    }

    # 加权随机算法
    $cumulativeWeights = @()
    $currentSum = 0
    foreach ($g in $groupList) {
        $currentSum += $g.Weight
        $cumulativeWeights += $currentSum
    }

    # 生成统计结果
    $statResults = 1..$frequency | ForEach-Object {
        $randomValue = $rnd.NextDouble() * $totalWeight
        $index = 0
        while ($index -lt $cumulativeWeights.Count - 1 -and $cumulativeWeights[$index] -lt $randomValue) {
            $index++
        }
        if ($index -lt $groupList.Count) {
            $groupList[$index].Number
        }
    } | Group-Object | ForEach-Object {
        [PSCustomObject]@{
            组号 = [int]$_.Name
            频数 = $_.Count
            占比 = [math]::Round(($_.Count / $frequency) * 100, 2)
        }
    } | Sort-Object -Property 频数 -Descending | Where-Object { $_.组号 -ne $null }

    # 输出统计结果
    Write-Host "`n各组的频数和占比统计：" -ForegroundColor Cyan
    $statResults | Format-Table -AutoSize -Property @(
        @{Label="组号"; Expression={$_.组号}},
        @{Label="频数"; Expression={$_.频数}; Alignment="Right"},
        @{Label="占比(%)"; Expression={$_.占比}; Alignment="Right"}
    )

    # 选择频数最高的组
    $selectedNumbers = @($statResults | Select-Object -First $groupNumber | ForEach-Object { $_.组号 })

    # 更新JSON状态
    foreach ($prop in $jsonContent.PSObject.Properties) {
        if ($prop.Name -ne 'Time') {
            $prop.Value.Selected = ($prop.Name -in $selectedNumbers)
        }
    }

    $jsonContent.Time = $currentDate

    # 保存更新后的JSON
    $jsonContent | ConvertTo-Json -Depth 3 | Set-Content $jsonPath -Encoding UTF8

    # 输出选中结果
    Write-Host "`n要交作业的组为:" -ForegroundColor Green
    if ($selectedNumbers.Count -gt 0) {
        $selectedNumbers | Sort-Object | ForEach-Object { 
            Write-Host "第${_}组"
        }
    } else {
        Write-Host "警告：未选中任何组" -ForegroundColor Yellow
    }

    # 输出成员名单
    $allMembers = @($selectedNumbers | ForEach-Object { 
        if ($jsonContent."$_") {
            $jsonContent."$_".Members 
        }
    } | Where-Object { $_ -ne $null } | Sort-Object -Culture "zh-CN" -Unique)
    
    Write-Host "`n要交作业的人员:" -ForegroundColor Green
    if ($allMembers.Count -gt 0) {
        Write-Host ($allMembers -join '、')
    } else {
        Write-Host "无有效人员信息" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "发生错误: $_" -ForegroundColor Red
    exit 1
}
finally {
    if ([Environment]::UserInteractive) {
        Write-Host "`n按任意键退出..." -NoNewline
        try { 
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') 
        } catch {
            Start-Sleep 3
        }
    }
}
