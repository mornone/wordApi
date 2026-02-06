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
        public bool IsRunning { get; private set; }
        public event Action<string>? OnLog;

        public WordService()
        {
            TaskDirectory = Path.Combine(AppContext.BaseDirectory, "Tasks");
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

        public async System.Threading.Tasks.Task StartAsync()
        {
            if (IsRunning) return;

            // 确保目录存在
            Directory.CreateDirectory(TaskDirectory);
            
            // 创建日志文件
            _logFilePath = Path.Combine(TaskDirectory, $"api_log_{DateTime.Now:yyyyMMdd}.txt");

            _cts = new CancellationTokenSource();
            IsRunning = true;

            Log($"服务启动 - 端口: {Port}, 任务目录: {TaskDirectory}");
            Log($"监听地址: http://0.0.0.0:{Port} (可通过局域网访问)");

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

                        // 保存上传的文件
                        var taskId = Guid.NewGuid().ToString();
                        var uploadDir = Path.Combine(TaskDirectory, "uploads");
                        Directory.CreateDirectory(uploadDir);
                        
                        var inputFilePath = Path.Combine(uploadDir, $"{taskId}_{file.FileName}");
                        Log($"POST /wordapi - {clientIp} - 保存文件到: {inputFilePath}");
                        
                        using (var stream = new FileStream(inputFilePath, FileMode.Create))
                        {
                            await file.CopyToAsync(stream);
                        }

                        Log($"POST /wordapi - {clientIp} - 文件保存成功");

                        var requestTask = new WordTask
                        {
                            TaskId = taskId,
                            InputFile = inputFilePath,
                            OutputDocx = Path.Combine(TaskDirectory, taskId + ".docx"),
                            OutputPdf = Path.Combine(TaskDirectory, taskId + ".pdf")
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
                        requestTask.OutputDocx = Path.Combine(TaskDirectory, requestTask.TaskId + ".docx");
                        requestTask.OutputPdf = Path.Combine(TaskDirectory, requestTask.TaskId + ".pdf");

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
            Log("ProcessQueue 线程已启动 (STA 模式)");
            
            while (!token.IsCancellationRequested)
            {
                if (_taskQueue.TryDequeue(out var task))
                {
                    _taskStatus[task.TaskId] = "running";
                    Log($"开始处理任务: {task.TaskId}, 输入文件: {task.InputFile}");
                    
                    try
                    {
                        Microsoft.Office.Interop.Word.Application word = new();
                        word.Visible = false;

                        Document doc;
                        
                        // 尝试使用受保护视图打开，如果失败则直接打开
                        try
                        {
                            var pv = word.ProtectedViewWindows.Open(task.InputFile);
                            doc = pv.Edit();
                            Log($"任务 {task.TaskId}: 从受保护视图打开文件");
                        }
                        catch
                        {
                            doc = word.Documents.Open(task.InputFile);
                            Log($"任务 {task.TaskId}: 直接打开文件");
                        }

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

                        doc.Close(false);
                        word.Quit();

                        _taskStatus[task.TaskId] = "completed";
                        var result = new Dictionary<string, string> { { "docx", task.OutputDocx } };
                        if (EnablePdf)
                            result["pdf"] = task.OutputPdf;
                        _taskResult[task.TaskId] = result;
                        
                        Log($"任务完成: {task.TaskId}");
                    }
                    catch (Exception ex)
                    {
                        _taskStatus[task.TaskId] = "failed";
                        _taskResult[task.TaskId] = new { error = ex.Message };
                        Log($"任务失败: {task.TaskId} - {ex.Message}");
                    }
                }
                else
                {
                    Thread.Sleep(200);
                }
            }
        }
    }
}
