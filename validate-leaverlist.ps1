# 文件路径
$filePath = ".\leaverlist.ps1"

# 读取原始内容
$lines = Get-Content $filePath

# 新文件内容（用于删除行）
$updatedLines = @()

# 使用 HashSet 避免重复
$newUsers = [System.Collections.Generic.HashSet[string]]::new()

foreach ($line in $lines) {

    $account = $line.Trim()

    # 保留空行
    if ([string]::IsNullOrWhiteSpace($account)) {
        $updatedLines += $line
        continue
    }

    # 如果不包含 "."，认为可能是组
    if ($account -notmatch "\.") {

        # 拼接成邮件地址
        $email = "$account@example.com"
        Write-Host "Processing possible group: $email"

        try {
            # 通过邮箱查找 AD 组
            $group = Get-ADGroup -Filter "mail -eq '$email'" -Properties mail -ErrorAction Stop

            if ($group) {
                Write-Host "  + Found group: $($group.Name)"

                $members = Get-ADGroupMember -Identity $group -Recursive | Where-Object {
                    $_.objectClass -eq "user"
                }

                foreach ($member in $members) {
                    $sam = $member.SamAccountName
                    if ($sam) {
                        if ($newUsers.Add($sam)) {
                            Write-Host "    - Added user: $sam"
                        }
                    }
                }
            } else {
                Write-Warning "  ! No AD group found with email $email"
            }

        } catch {
            Write-Warning "Failed processing group: $email. $_"
        }

        # 删除原组行
        continue
    }

    # 普通用户保留
    $updatedLines += $line
}

# 写回文件（已删除组行）
Set-Content -Path $filePath -Value $updatedLines -Encoding UTF8

# 追加组成员到末尾
if ($newUsers.Count -gt 0) {
    $usersToAdd = $newUsers | Sort-Object
    Add-Content -Path $filePath -Value $usersToAdd
}

# --- 最终文件去重（验证并删除重复行） ---
$allLines = Get-Content $filePath
$uniqueLines = $allLines | Select-Object -Unique
Set-Content -Path $filePath -Value $uniqueLines -Encoding UTF8
