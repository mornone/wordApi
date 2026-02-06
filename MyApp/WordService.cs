using System.Collections.Concurrent;
using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
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
        
        public bool EnableRefresh { get; set; } = true;
        public bool EnablePdf { get; set; } = true;
        public int Port { get; set; } = 5000;
        public string TaskDirectory { get; set; } = string.Empty;
        public bool IsRunning { get; private set; }

        public WordService()
        {
            TaskDirectory = Path.Combine(AppContext.BaseDirectory, "Tasks");
        }

        public async System.Threading.Tasks.Task StartAsync()
        {
            if (IsRunning) return;

            // 确保目录存在
            Directory.CreateDirectory(TaskDirectory);

            _cts = new CancellationTokenSource();
            IsRunning = true;

            // 启动后台处理线程
            _ = System.Threading.Tasks.Task.Run(() => ProcessQueue(_cts.Token));

            // 启动 HTTP 服务
            var builder = WebApplication.CreateBuilder();
            builder.WebHost.UseUrls($"http://localhost:{Port}");
            
            _app = builder.Build();

            _app.MapPost("/wordapi", async (HttpContext context) =>
            {
                var requestTask = await JsonSerializer.DeserializeAsync<WordTask>(context.Request.Body);
                if (requestTask == null || !File.Exists(requestTask.InputFile))
                {
                    context.Response.StatusCode = 400;
                    await context.Response.WriteAsync("InputFile does not exist.");
                    return;
                }

                requestTask.TaskId = Guid.NewGuid().ToString();
                requestTask.OutputDocx = Path.Combine(TaskDirectory, requestTask.TaskId + ".docx");
                requestTask.OutputPdf = Path.Combine(TaskDirectory, requestTask.TaskId + ".pdf");

                _taskQueue.Enqueue(requestTask);
                _taskStatus[requestTask.TaskId] = "queued";

                await context.Response.WriteAsJsonAsync(new { status = "queued", taskId = requestTask.TaskId });
            });

            _app.MapGet("/wordapi", async (HttpContext context) =>
            {
                var taskId = context.Request.Query["taskId"].ToString() ?? string.Empty;
                if (string.IsNullOrEmpty(taskId))
                {
                    context.Response.StatusCode = 400;
                    await context.Response.WriteAsync("Missing taskId query.");
                    return;
                }

                if (_taskStatus.ContainsKey(taskId))
                {
                    _taskResult.TryGetValue(taskId, out var result);
                    await context.Response.WriteAsJsonAsync(new { taskId, status = _taskStatus[taskId], result });
                }
                else
                {
                    context.Response.StatusCode = 404;
                    await context.Response.WriteAsync("taskId not found.");
                }
            });

            // 在后台线程启动 Web 服务
            _webHostTask = System.Threading.Tasks.Task.Run(async () => await _app.RunAsync());
            
            // 等待一小段时间确保服务启动
            await System.Threading.Tasks.Task.Delay(500);
        }

        public async System.Threading.Tasks.Task StopAsync()
        {
            if (!IsRunning) return;

            IsRunning = false;
            
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
            }
            catch
            {
                // 忽略停止时的异常
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
            while (!token.IsCancellationRequested)
            {
                if (_taskQueue.TryDequeue(out var task))
                {
                    _taskStatus[task.TaskId] = "running";
                    try
                    {
                        Microsoft.Office.Interop.Word.Application word = new();
                        word.Visible = false;

                        var pv = word.ProtectedViewWindows.Open(task.InputFile);
                        var doc = pv.Edit();

                        if (EnableRefresh)
                        {
                            doc.Fields.Update();
                            foreach (TableOfContents toc in doc.TablesOfContents)
                                toc.Update();
                        }

                        doc.SaveAs2(task.OutputDocx);

                        if (EnablePdf)
                        {
                            doc.ExportAsFixedFormat(task.OutputPdf, WdExportFormat.wdExportFormatPDF);
                        }

                        doc.Close(false);
                        word.Quit();

                        _taskStatus[task.TaskId] = "completed";
                        var result = new Dictionary<string, string> { { "docx", task.OutputDocx } };
                        if (EnablePdf)
                            result["pdf"] = task.OutputPdf;
                        _taskResult[task.TaskId] = result;
                    }
                    catch (Exception ex)
                    {
                        _taskStatus[task.TaskId] = "failed";
                        _taskResult[task.TaskId] = new { error = ex.Message };
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
