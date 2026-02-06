# Word API 服务使用说明

## 服务已正常运行 ✅

如果你看到 "taskId not found" 响应，说明 API 服务已经成功启动！

## API 接口说明

### 1. 提交任务（POST）

**请求地址：** `http://localhost:5000/wordapi`

**请求方法：** POST

**请求体（JSON）：**
```json
{
  "InputFile": "C:\\path\\to\\your\\document.docx"
}
```

**响应示例：**
```json
{
  "status": "queued",
  "taskId": "abc123-def456-789..."
}
```

### 2. 查询任务状态（GET）

**请求地址：** `http://localhost:5000/wordapi?taskId=YOUR_TASK_ID`

**请求方法：** GET

**响应示例（处理中）：**
```json
{
  "taskId": "abc123-def456-789...",
  "status": "running",
  "result": null
}
```

**响应示例（完成）：**
```json
{
  "taskId": "abc123-def456-789...",
  "status": "completed",
  "result": {
    "docx": "E:\\path\\to\\output\\abc123-def456-789.docx",
    "pdf": "E:\\path\\to\\output\\abc123-def456-789.pdf"
  }
}
```

**响应示例（失败）：**
```json
{
  "taskId": "abc123-def456-789...",
  "status": "failed",
  "result": {
    "error": "错误信息"
  }
}
```

## 使用 PowerShell 测试

### 提交任务
```powershell
$body = @{
    InputFile = "C:\test\document.docx"
} | ConvertTo-Json

$response = Invoke-RestMethod -Uri "http://localhost:5000/wordapi" -Method Post -Body $body -ContentType "application/json"
$taskId = $response.taskId
Write-Host "任务ID: $taskId"
```

### 查询状态
```powershell
$taskId = "YOUR_TASK_ID_HERE"
$status = Invoke-RestMethod -Uri "http://localhost:5000/wordapi?taskId=$taskId"
$status | ConvertTo-Json
```

### 完整流程示例
```powershell
# 1. 提交任务
$body = @{
    InputFile = "C:\test\document.docx"
} | ConvertTo-Json

$response = Invoke-RestMethod -Uri "http://localhost:5000/wordapi" -Method Post -Body $body -ContentType "application/json"
$taskId = $response.taskId
Write-Host "任务已提交，ID: $taskId"

# 2. 等待并查询状态
Start-Sleep -Seconds 2

$status = Invoke-RestMethod -Uri "http://localhost:5000/wordapi?taskId=$taskId"
Write-Host "任务状态: $($status.status)"

if ($status.status -eq "completed") {
    Write-Host "DOCX 文件: $($status.result.docx)"
    Write-Host "PDF 文件: $($status.result.pdf)"
}
```

## 使用 curl 测试

### 提交任务
```bash
curl -X POST http://localhost:5000/wordapi ^
  -H "Content-Type: application/json" ^
  -d "{\"InputFile\":\"C:\\test\\document.docx\"}"
```

### 查询状态
```bash
curl "http://localhost:5000/wordapi?taskId=YOUR_TASK_ID"
```

## 使用浏览器测试

直接在浏览器访问（仅限 GET 请求）：
```
http://localhost:5000/wordapi?taskId=YOUR_TASK_ID
```

## 任务状态说明

- **queued** - 任务已加入队列，等待处理
- **running** - 任务正在处理中
- **completed** - 任务处理完成
- **failed** - 任务处理失败

## 注意事项

1. ✅ InputFile 必须是完整的文件路径
2. ✅ 文件必须存在且可访问
3. ✅ 运行服务的电脑必须安装 Microsoft Office（Word）
4. ✅ 任务是异步处理的，提交后需要轮询查询状态
5. ✅ 输出文件保存在配置的任务目录中

## 功能配置

在 UI 界面中可以配置：
- **服务端口** - 默认 5000
- **任务目录** - 输出文件保存位置
- **启用目录刷新** - 是否更新文档中的域和目录
- **启用 PDF 转换** - 是否同时生成 PDF 文件
- **开机自动启动** - 是否随 Windows 启动

## 故障排查

### API 无法访问
- 检查服务是否启动（界面显示"运行中"）
- 检查端口是否被占用
- 尝试更换端口号

### 任务一直处于 queued 状态
- 检查后台处理线程是否正常
- 重启服务

### 任务失败
- 检查 Word 文档是否损坏
- 检查是否安装了 Microsoft Office
- 查看 result.error 中的错误信息
