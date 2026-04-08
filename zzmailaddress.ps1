# 读取文件内容
$filePath = ".\leaverlist.ps1"
$fileContent = Get-Content $filePath

# 使用正则表达式提取邮箱地址
$emailAddresses = [Regex]::Matches($fileContent, "\b[\w\.-]+@[\w\.-]+\.\w{2,4}\b")

# 去除所有邮箱后缀（@域名.后缀）
$cleanedEmailAddresses = $emailAddresses | ForEach-Object {
    $_.Value -replace "@[\w\.-]+\.\w{2,4}", ""
}

# 将结果覆盖到文件，使用换行符分隔
Set-Content -Path $filePath -Value ($cleanedEmailAddresses -join "`n")

.\validate-leaverlist.ps1
