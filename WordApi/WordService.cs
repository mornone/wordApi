using System.Collections.Concurrent;
using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Word;

namespace WordApiService
{
    public class WordService
    {
        private WebApplication? _app;
        private System.Threading.Tasks.Task? _webHostTask;
        private CancellationTokenSource? _cts;
        private readonly ConcurrentQueue<WordTask> _taskQueue = new();
        private readonly ConcurrentDictionary<string, string> _taskStatus = new();
        private readonly ConcurrentDictionary<string, object> _taskResult = new();
        private string _logFilePath = string.Empty;
        
        public bool EnableRefresh { get; set; } = true;
        public bool EnablePdf { get; set; } = true;
        public int Port { get; set; } = 5000;
        public string TaskDirectory { get; set; } = string.Empty;
        public string UploadDirectory { get; set; } = string.Empty;
        public string OutputDirectory { get; set; } = string.Empty;
        public bool AutoDeleteUploads { get; set; } = false;
        public bool AutoDeleteOutputs { get; set; } = false;
        public int DeleteAfterDays { get; set; } = 7;
        public bool IsRunning { get; private set; }
        public event Action<string>? OnLog;

        public WordService()
        {
            TaskDirectory = Path.Combine(AppContext.BaseDirectory, "Tasks");
            UploadDirectory = Path.Combine(TaskDirectory, "uploads");
            OutputDirectory = Path.Combine(TaskDirectory, "outputs");
        }

        private void Log(string message)
        {
            var logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            
            // 触发事件通知界面
            OnLog?.Invoke(logMessage);
            
            // 写入日志文件
            try
            {
                File.AppendAllText(_logFilePath, logMessage + Environment.NewLine);
            }
            catch
            {
                // 忽略日志写入错误
            }
        }

        private string GetFileUrl(string filePath)
        {
            // 将文件路径转换为相对 URL 路径
            // 例如: E:\...\outputs\2026\0206\xxx.docx -> /files/2026/0206/xxx.docx
            try
            {
                var relativePath = Path.GetRelativePath(OutputDirectory, filePath);
                var parts = relativePath.Split(Path.DirectorySeparatorChar);
                if (parts.Length >= 3)
                {
                    var year = parts[0];
                    var day = parts[1];
                    var filename = parts[2];
                    return $"/files/{year}/{day}/{filename}";
                }
            }
            catch
            {
                // 如果转换失败，返回原路径
            }
            return filePath;
        }

        private string GetFallbackHtml()
        {
            return @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <title>Word API 服务</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 800px; margin: 50px auto; padding: 20px; }
        h1 { color: #667eea; }
        .endpoint { background: #f5f5f5; padding: 15px; margin: 10px 0; border-radius: 5px; }
    </style>
</head>
<body>
    <h1>Word API 服务</h1>
    <p>欢迎使用 Word API 服务！</p>
    <div class='endpoint'>
        <h3>POST /wordapi</h3>
        <p>上传并处理 Word 文档（multipart/form-data，字段名：InputFile）</p>
    </div>
    <div class='endpoint'>
        <h3>GET /wordapi?taskId={taskId}</h3>
        <p>查询任务状态和结果</p>
    </div>
    <div class='endpoint'>
        <h3>GET /files/{year}/{day}/{filename}</h3>
        <p>下载处理后的文件</p>
    </div>
    <p style='margin-top: 30px; color: #666;'>© 2026 信安世纪 - www.secdriver.com</p>
</body>
</html>";
        }

        private void CleanOldFiles()
        {
            try
            {
                var cutoffDate = DateTime.Now.AddDays(-DeleteAfterDays);
                
                if (AutoDeleteUploads && Directory.Exists(UploadDirectory))
                {
                    var deletedCount = 0;
                    foreach (var yearDir in Directory.GetDirectories(UploadDirectory))
                    {
                        foreach (var dayDir in Directory.GetDirectories(yearDir))
                        {
                            var dirInfo = new DirectoryInfo(dayDir);
                            if (dirInfo.CreationTime < cutoffDate)
                            {
                                Directory.Delete(dayDir, true);
                                deletedCount++;
                            }
                        }
                        // 删除空的年份目录
                        if (Directory.GetDirectories(yearDir).Length == 0)
                        {
                            Directory.Delete(yearDir);
                        }
                    }
                    if (deletedCount > 0)
                        Log($"清理上传文件: 删除了 {deletedCount} 个超过 {DeleteAfterDays} 天的目录");
                }
                
                if (AutoDeleteOutputs && Directory.Exists(OutputDirectory))
                {
                    var deletedCount = 0;
                    foreach (var yearDir in Directory.GetDirectories(OutputDirectory))
                    {
                        foreach (var dayDir in Directory.GetDirectories(yearDir))
                        {
                            var dirInfo = new DirectoryInfo(dayDir);
                            if (dirInfo.CreationTime < cutoffDate)
                            {
                                Directory.Delete(dayDir, true);
                                deletedCount++;
                            }
                        }
                        // 删除空的年份目录
                        if (Directory.GetDirectories(yearDir).Length == 0)
                        {
                            Directory.Delete(yearDir);
                        }
                    }
                    if (deletedCount > 0)
                        Log($"清理输出文件: 删除了 {deletedCount} 个超过 {DeleteAfterDays} 天的目录");
                }
            }
            catch (Exception ex)
            {
                Log($"清理旧文件时出错: {ex.Message}");
            }
        }

        public async System.Threading.Tasks.Task StartAsync()
        {
            if (IsRunning) return;

            // 确保目录存在
            Directory.CreateDirectory(TaskDirectory);
            Directory.CreateDirectory(UploadDirectory);
            Directory.CreateDirectory(OutputDirectory);
            
            // 创建日志文件
            _logFilePath = Path.Combine(TaskDirectory, $"api_log_{DateTime.Now:yyyyMMdd}.txt");

            _cts = new CancellationTokenSource();
            IsRunning = true;

            Log($"服务启动 - 端口: {Port}, 任务目录: {TaskDirectory}");
            Log($"监听地址: http://0.0.0.0:{Port} (可通过局域网访问)");
            Log($"上传目录: {UploadDirectory}");
            Log($"输出目录: {OutputDirectory}");
            
            // 清理旧文件
            if (AutoDeleteUploads || AutoDeleteOutputs)
            {
                Log($"文件清理策略: 上传文件={AutoDeleteUploads}, 输出文件={AutoDeleteOutputs}, 保留天数={DeleteAfterDays}");
                CleanOldFiles();
            }

            // 启动后台处理线程（必须使用 STA 线程，Word COM 需要）
            Log("准备启动 ProcessQueue 后台线程...");
            var processThread = new Thread(() =>
            {
                try
                {
                    ProcessQueue(_cts.Token);
                }
                catch (Exception ex)
                {
                    Log($"ProcessQueue 线程异常: {ex.Message}");
                }
            });
            processThread.SetApartmentState(ApartmentState.STA);
            processThread.IsBackground = true;
            processThread.Start();
            Log("ProcessQueue 后台线程已提交启动");

            // 启动 HTTP 服务
            var args = new string[] { };
            var builder = WebApplication.CreateBuilder(args);
            builder.WebHost.UseUrls($"http://0.0.0.0:{Port}");  // 监听所有网络接口
            builder.Logging.ClearProviders(); // 清除日志输出
            
            _app = builder.Build();

            _app.MapPost("/wordapi", async (HttpContext context) =>
            {
                var clientIp = context.Connection.RemoteIpAddress?.ToString() ?? "unknown";
                Log($"POST /wordapi - {clientIp} - 收到请求, ContentType: {context.Request.ContentType}");
                
                try
                {
                    // 检查是否是文件上传请求
                    if (context.Request.HasFormContentType && context.Request.Form.Files.Count > 0)
                    {
                        Log($"POST /wordapi - {clientIp} - 检测到文件上传, 文件数量: {context.Request.Form.Files.Count}");
                        
                        var file = context.Request.Form.Files["InputFile"];
                        if (file == null || file.Length == 0)
                        {
                            Log($"POST /wordapi - {clientIp} - 400 Bad Request: InputFile is required");
                            context.Response.StatusCode = 400;
                            await context.Response.WriteAsJsonAsync(new { error = "InputFile is required." });
                            return;
                        }

                        Log($"POST /wordapi - {clientIp} - 文件名: {file.FileName}, 大小: {file.Length} bytes");

                        // 保存上传的文件，使用 年/月日 目录结构
                        var taskId = Guid.NewGuid().ToString();
                        var now = DateTime.Now;
                        var uploadDir = Path.Combine(UploadDirectory, now.ToString("yyyy"), now.ToString("MMdd"));
                        Directory.CreateDirectory(uploadDir);
                        
                        var inputFilePath = Path.Combine(uploadDir, $"{taskId}_{file.FileName}");
                        Log($"POST /wordapi - {clientIp} - 保存文件到: {inputFilePath}");
                        
                        using (var stream = new FileStream(inputFilePath, FileMode.Create))
                        {
                            await file.CopyToAsync(stream);
                        }

                        Log($"POST /wordapi - {clientIp} - 文件保存成功");

                        // 输出文件也使用 年/月日 目录结构
                        var outputDir = Path.Combine(OutputDirectory, now.ToString("yyyy"), now.ToString("MMdd"));
                        Directory.CreateDirectory(outputDir);

                        var requestTask = new WordTask
                        {
                            TaskId = taskId,
                            InputFile = inputFilePath,
                            OutputDocx = Path.Combine(outputDir, taskId + ".docx"),
                            OutputPdf = Path.Combine(outputDir, taskId + ".pdf")
                        };

                        _taskQueue.Enqueue(requestTask);
                        _taskStatus[requestTask.TaskId] = "queued";

                        Log($"POST /wordapi - {clientIp} - 200 OK: Task created - TaskId: {requestTask.TaskId}, Input: {file.FileName}, 队列长度: {_taskQueue.Count}");
                        await context.Response.WriteAsJsonAsync(new { status = "queued", taskId = requestTask.TaskId });
                    }
                    else
                    {
                        // 原有的 JSON 请求处理
                        var requestTask = await JsonSerializer.DeserializeAsync<WordTask>(context.Request.Body);
                        if (requestTask == null || string.IsNullOrEmpty(requestTask.InputFile))
                        {
                            Log($"POST /wordapi - {clientIp} - 400 Bad Request: InputFile is required");
                            context.Response.StatusCode = 400;
                            await context.Response.WriteAsJsonAsync(new { error = "InputFile is required." });
                            return;
                        }

                        if (!File.Exists(requestTask.InputFile))
                        {
                            Log($"POST /wordapi - {clientIp} - 400 Bad Request: File not found - {requestTask.InputFile}");
                            context.Response.StatusCode = 400;
                            await context.Response.WriteAsJsonAsync(new { error = "InputFile does not exist." });
                            return;
                        }

                        requestTask.TaskId = Guid.NewGuid().ToString();
                        
                        // 输出文件使用 年/月日 目录结构
                        var now = DateTime.Now;
                        var outputDir = Path.Combine(OutputDirectory, now.ToString("yyyy"), now.ToString("MMdd"));
                        Directory.CreateDirectory(outputDir);
                        
                        requestTask.OutputDocx = Path.Combine(outputDir, requestTask.TaskId + ".docx");
                        requestTask.OutputPdf = Path.Combine(outputDir, requestTask.TaskId + ".pdf");

                        _taskQueue.Enqueue(requestTask);
                        _taskStatus[requestTask.TaskId] = "queued";

                        Log($"POST /wordapi - {clientIp} - 200 OK: Task created - TaskId: {requestTask.TaskId}, Input: {requestTask.InputFile}");
                        await context.Response.WriteAsJsonAsync(new { status = "queued", taskId = requestTask.TaskId });
                    }
                }
                catch (Exception ex)
                {
                    Log($"POST /wordapi - {clientIp} - 500 Error: {ex.Message}");
                    context.Response.StatusCode = 500;
                    await context.Response.WriteAsJsonAsync(new { error = "Internal server error." });
                }
            });

            _app.MapGet("/wordapi", async (HttpContext context) =>
            {
                var clientIp = context.Connection.RemoteIpAddress?.ToString() ?? "unknown";
                var taskId = context.Request.Query["taskId"].ToString() ?? string.Empty;
                
                if (string.IsNullOrEmpty(taskId))
                {
                    Log($"GET /wordapi - {clientIp} - 400 Bad Request: Missing taskId");
                    context.Response.StatusCode = 400;
                    await context.Response.WriteAsJsonAsync(new { error = "Missing taskId query parameter." });
                    return;
                }

                if (_taskStatus.ContainsKey(taskId))
                {
                    _taskResult.TryGetValue(taskId, out var result);
                    var status = _taskStatus[taskId];
                    Log($"GET /wordapi - {clientIp} - 200 OK: TaskId: {taskId}, Status: {status}");
                    await context.Response.WriteAsJsonAsync(new { taskId, status, result });
                }
                else
                {
                    Log($"GET /wordapi - {clientIp} - 404 Not Found: TaskId: {taskId}");
                    context.Response.StatusCode = 404;
                    await context.Response.WriteAsJsonAsync(new { error = "taskId not found." });
                }
            });

            // API 文档端点 - 根路径
            _app.MapGet("/", async (HttpContext context) =>
            {
                var clientIp = context.Connection.RemoteIpAddress?.ToString() ?? "unknown";
                var htmlPath = Path.Combine(AppContext.BaseDirectory, "api-docs.html");
                
                if (File.Exists(htmlPath))
                {
                    Log($"GET / - {clientIp} - 200 OK: API 文档");
                    context.Response.ContentType = "text/html; charset=utf-8";
                    await context.Response.SendFileAsync(htmlPath);
                }
                else
                {
                    Log($"GET / - {clientIp} - 200 OK: 备用文档");
                    context.Response.ContentType = "text/html; charset=utf-8";
                    await context.Response.WriteAsync(GetFallbackHtml());
                }
            });

            // API 文档端点 - /docs 路径
            _app.MapGet("/docs", async (HttpContext context) =>
            {
                var clientIp = context.Connection.RemoteIpAddress?.ToString() ?? "unknown";
                var htmlPath = Path.Combine(AppContext.BaseDirectory, "api-docs.html");
                
                if (File.Exists(htmlPath))
                {
                    Log($"GET /docs - {clientIp} - 200 OK: API 文档");
                    context.Response.ContentType = "text/html; charset=utf-8";
                    await context.Response.SendFileAsync(htmlPath);
                }
                else
                {
                    Log($"GET /docs - {clientIp} - 200 OK: 备用文档");
                    context.Response.ContentType = "text/html; charset=utf-8";
                    await context.Response.WriteAsync(GetFallbackHtml());
                }
            });

            // 静态文件服务 - 提供输出文件下载
            _app.MapGet("/files/{year}/{day}/{filename}", async (HttpContext context, string year, string day, string filename) =>
            {
                var clientIp = context.Connection.RemoteIpAddress?.ToString() ?? "unknown";
                var filePath = Path.Combine(OutputDirectory, year, day, filename);
                
                if (!File.Exists(filePath))
                {
                    Log($"GET /files/{year}/{day}/{filename} - {clientIp} - 404 Not Found");
                    context.Response.StatusCode = 404;
                    await context.Response.WriteAsJsonAsync(new { error = "File not found." });
                    return;
                }

                Log($"GET /files/{year}/{day}/{filename} - {clientIp} - 200 OK");
                
                // 设置正确的 Content-Type
                var extension = Path.GetExtension(filename).ToLower();
                context.Response.ContentType = extension switch
                {
                    ".pdf" => "application/pdf",
                    ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    _ => "application/octet-stream"
                };
                
                // 设置文件名
                context.Response.Headers.Append("Content-Disposition", $"attachment; filename=\"{filename}\"");
                
                await context.Response.SendFileAsync(filePath);
            });

            // 在后台线程启动 Web 服务
            _webHostTask = _app.StartAsync();
            
            // 等待服务启动
            await System.Threading.Tasks.Task.Delay(1000);
            Log("HTTP 服务已启动");
        }

        public async System.Threading.Tasks.Task StopAsync()
        {
            if (!IsRunning) return;

            IsRunning = false;
            Log("服务停止中...");
            
            try
            {
                _cts?.Cancel();
                
                if (_app != null)
                {
                    await _app.StopAsync();
                    await _app.DisposeAsync();
                }

                if (_webHostTask != null)
                {
                    await System.Threading.Tasks.Task.WhenAny(_webHostTask, System.Threading.Tasks.Task.Delay(2000));
                }
                
                Log("服务已停止");
            }
            catch (Exception ex)
            {
                Log($"停止服务时出错: {ex.Message}");
            }
            finally
            {
                _cts?.Dispose();
                _cts = null;
                _app = null;
                _webHostTask = null;
            }
        }

        private void ProcessQueue(CancellationToken token)
        {
            Log($"ProcessQueue 线程已启动 (STA 模式: {Thread.CurrentThread.GetApartmentState() == ApartmentState.STA})");
            
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                Log("警告: 线程不是 STA 模式，Word COM 可能无法正常工作");
            }
            
            while (!token.IsCancellationRequested)
            {
                if (_taskQueue.TryDequeue(out var task))
                {
                    _taskStatus[task.TaskId] = "running";
                    Log($"开始处理任务: {task.TaskId}, 输入文件: {task.InputFile}");
                    
                    Microsoft.Office.Interop.Word.Application? word = null;
                    Document? doc = null;
                    
                    try
                    {
                        // 创建 Word 应用程序实例（带重试）
                        Log($"任务 {task.TaskId}: 创建 Word 应用程序实例");
                        
                        int retryCount = 0;
                        int maxRetries = 3;
                        Exception? lastException = null;
                        
                        while (retryCount < maxRetries && word == null)
                        {
                            try
                            {
                                word = new Microsoft.Office.Interop.Word.Application();
                                word.Visible = false;
                                word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                                Log($"任务 {task.TaskId}: Word 应用程序已创建");
                            }
                            catch (Exception ex)
                            {
                                lastException = ex;
                                retryCount++;
                                if (retryCount < maxRetries)
                                {
                                    Log($"任务 {task.TaskId}: Word 创建失败，重试 {retryCount}/{maxRetries}");
                                    Thread.Sleep(1000);
                                }
                            }
                        }
                        
                        if (word == null)
                        {
                            throw new Exception($"无法创建 Word 应用程序实例: {lastException?.Message}");
                        }

                        // 直接打开文件，不使用受保护视图
                        Log($"任务 {task.TaskId}: 打开文件 - {Path.GetFileName(task.InputFile)}");
                        
                        // 使用更安全的打开方式
                        object fileName = task.InputFile;
                        object confirmConversions = false;
                        object readOnly = false;
                        object addToRecentFiles = false;
                        object visible = false;
                        object missing = System.Reflection.Missing.Value;
                        
                        doc = word.Documents.Open(
                            ref fileName,
                            ref confirmConversions,
                            ref readOnly,
                            ref addToRecentFiles,
                            ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref visible,
                            ref missing, ref missing, ref missing, ref missing
                        );
                        
                        Log($"任务 {task.TaskId}: 文件打开成功");

                        if (EnableRefresh)
                        {
                            Log($"任务 {task.TaskId}: 更新域和目录");
                            doc.Fields.Update();
                            foreach (TableOfContents toc in doc.TablesOfContents)
                                toc.Update();
                        }
                        else
                        {
                            Log($"任务 {task.TaskId}: 跳过目录刷新（配置已禁用）");
                        }

                        Log($"任务 {task.TaskId}: 保存 DOCX");
                        doc.SaveAs2(task.OutputDocx);

                        if (EnablePdf)
                        {
                            Log($"任务 {task.TaskId}: 导出 PDF");
                            doc.ExportAsFixedFormat(task.OutputPdf, WdExportFormat.wdExportFormatPDF);
                        }
                        else
                        {
                            Log($"任务 {task.TaskId}: 跳过 PDF 导出（配置已禁用）");
                        }

                        Log($"任务 {task.TaskId}: 关闭文档");
                        doc.Close(false);
                        
                        // 释放文档对象
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                        doc = null;
                        
                        Log($"任务 {task.TaskId}: 退出 Word 应用程序");
                        word.Quit();
                        
                        // 释放 Word 应用程序对象
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(word);
                        word = null;
                        
                        // 强制垃圾回收，确保 COM 对象被释放
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();

                        _taskStatus[task.TaskId] = "completed";
                        
                        // 返回 HTTP URL 而不是文件路径
                        var result = new Dictionary<string, string> { { "docx", GetFileUrl(task.OutputDocx) } };
                        if (EnablePdf)
                            result["pdf"] = GetFileUrl(task.OutputPdf);
                        _taskResult[task.TaskId] = result;
                        
                        Log($"任务完成: {task.TaskId}");
                        Log("----------------------------------------");
                    }
                    catch (Exception ex)
                    {
                        _taskStatus[task.TaskId] = "failed";
                        _taskResult[task.TaskId] = new { error = ex.Message };
                        Log($"任务失败: {task.TaskId} - {ex.Message}");
                        Log($"错误详情: {ex.GetType().Name} - {ex.StackTrace}");
                        Log("----------------------------------------");
                        
                        // 确保清理资源
                        try
                        {
                            Log($"任务 {task.TaskId}: 清理资源");
                            
                            if (doc != null)
                            {
                                try
                                {
                                    doc.Close(false);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                                }
                                catch (Exception cleanEx)
                                {
                                    Log($"任务 {task.TaskId}: 清理文档对象失败 - {cleanEx.Message}");
                                }
                                doc = null;
                            }
                            
                            if (word != null)
                            {
                                try
                                {
                                    word.Quit();
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(word);
                                }
                                catch (Exception cleanEx)
                                {
                                    Log($"任务 {task.TaskId}: 清理 Word 应用程序失败 - {cleanEx.Message}");
                                }
                                word = null;
                            }
                            
                            // 强制垃圾回收
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                            
                            Log($"任务 {task.TaskId}: 资源清理完成");
                        }
                        catch (Exception cleanupEx)
                        {
                            Log($"任务 {task.TaskId}: 资源清理异常 - {cleanupEx.Message}");
                        }
                    }
                }
                else
                {
                    Thread.Sleep(200);
                }
                
                // 任务之间添加短暂延迟，确保资源完全释放
                Thread.Sleep(500);
            }
        }
    }
}
