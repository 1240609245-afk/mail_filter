cd "$env:USERPROFILE\Desktop\mail_filter"

Write-Host "开始运行邮件检测脚本..."

python mail_filter.py

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ 脚本运行失败，停止上传"
    exit 1
}

Write-Host "脚本运行完成，开始上传..."

git add .

git diff --cached --quiet
if ($LASTEXITCODE -ne 0) {
    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    git commit -m "Auto update $time"
    git push
    Write-Host "✅ 上传成功，网页已更新"
} else {
    Write-Host "⚠️ 没有变化，不需要上传"
}