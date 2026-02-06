using System.Windows.Forms;
using Microsoft.Win32;

namespace WordApiService
{
    public class MainForm : Form
    {
        private readonly WordService _service;
        private NumericUpDown _portInput = null!;
        private CheckBox _refreshCheckBox = null!;
        private CheckBox _pdfCheckBox = null!;
        private CheckBox _autoStartCheckBox = null!;
        private TextBox _taskDirInput = null!;
        private Button _browseDirButton = null!;
        private Button _startButton = null!;
        private Label _statusLabel = null!;
        private TextBox _logTextBox = null!;

        private const string AutoStartRegKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string AppName = "WordApiService";

        public MainForm()
        {
            _service = new WordService();
            _service.OnLog += OnServiceLog;
            InitializeUI();
            LoadAutoStartStatus();
            
            // 调试：显示嵌入的资源名称
            LogEmbeddedResources();
        }

        private void LogEmbeddedResources()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var resources = assembly.GetManifestResourceNames();
                if (resources.Length > 0)
                {
                    OnServiceLog("嵌入的资源列表:");
                    foreach (var resource in resources)
                    {
                        OnServiceLog($"  - {resource}");
                    }
                }
            }
            catch
            {
                // 忽略错误
            }
        }

        private void OnServiceLog(string message)
        {
            if (_logTextBox.InvokeRequired)
            {
                _logTextBox.Invoke(() => OnServiceLog(message));
                return;
            }

            _logTextBox.AppendText(message + Environment.NewLine);
            _logTextBox.SelectionStart = _logTextBox.Text.Length;
            _logTextBox.ScrollToCaret();
        }

        private void InitializeUI()
        {
            Text = "Word API 服务";
            Width = 450;
            Height = 550;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = true;

            // 设置窗口图标（从嵌入资源加载）
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                
                // 尝试多个可能的资源名称
                string[] possibleNames = new[]
                {
                    "WordApiService.icon.ico",
                    "MyApp.icon.ico",
                    "icon.ico"
                };

                foreach (var resourceName in possibleNames)
                {
                    using var stream = assembly.GetManifestResourceStream(resourceName);
                    if (stream != null)
                    {
                        Icon = new Icon(stream);
                        break;
                    }
                }
            }
            catch
            {
                // 如果加载图标失败，使用默认图标
            }

            // 端口配置
            var portLabel = new Label
            {
                Text = "服务端口:",
                Left = 20,
                Top = 20,
                Width = 80
            };
            Controls.Add(portLabel);

            _portInput = new NumericUpDown
            {
                Left = 110,
                Top = 20,
                Width = 300,
                Minimum = 1000,
                Maximum = 65535,
                Value = 5000
            };
            Controls.Add(_portInput);

            // 任务目录配置
            var taskDirLabel = new Label
            {
                Text = "任务目录:",
                Left = 20,
                Top = 55,
                Width = 80
            };
            Controls.Add(taskDirLabel);

            _taskDirInput = new TextBox
            {
                Left = 110,
                Top = 55,
                Width = 240,
                Text = Path.Combine(AppContext.BaseDirectory, "Tasks")
            };
            Controls.Add(_taskDirInput);

            _browseDirButton = new Button
            {
                Text = "浏览",
                Left = 360,
                Top = 53,
                Width = 50,
                Height = 25
            };
            _browseDirButton.Click += BrowseDirButton_Click;
            Controls.Add(_browseDirButton);

            // 刷新目录开关
            _refreshCheckBox = new CheckBox
            {
                Text = "启用目录刷新（更新域和目录）",
                Left = 20,
                Top = 95,
                Width = 390,
                Checked = true
            };
            Controls.Add(_refreshCheckBox);

            // 转PDF开关
            _pdfCheckBox = new CheckBox
            {
                Text = "启用 PDF 转换",
                Left = 20,
                Top = 125,
                Width = 390,
                Checked = true
            };
            Controls.Add(_pdfCheckBox);

            // 开机自启开关
            _autoStartCheckBox = new CheckBox
            {
                Text = "开机自动启动",
                Left = 20,
                Top = 155,
                Width = 390,
                Checked = false
            };
            _autoStartCheckBox.CheckedChanged += AutoStartCheckBox_CheckedChanged;
            Controls.Add(_autoStartCheckBox);

            // 启动按钮
            _startButton = new Button
            {
                Text = "启动服务",
                Left = 20,
                Top = 195,
                Width = 390,
                Height = 40
            };
            _startButton.Click += StartButton_Click;
            Controls.Add(_startButton);

            // 状态标签
            _statusLabel = new Label
            {
                Text = "状态: 未启动",
                Left = 20,
                Top = 245,
                Width = 590,
                Height = 30,
                ForeColor = Color.Gray
            };
            Controls.Add(_statusLabel);

            // 日志显示区域
            var logLabel = new Label
            {
                Text = "运行日志:",
                Left = 20,
                Top = 285,
                Width = 100,
                Height = 20
            };
            Controls.Add(logLabel);

            _logTextBox = new TextBox
            {
                Left = 20,
                Top = 310,
                Width = 390,
                Height = 150,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                BackColor = Color.White,
                Font = new Font("Consolas", 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            Controls.Add(_logTextBox);

            // 版权信息
            var copyrightLabel = new Label
            {
                Text = "---------- www.secdriver.com 信安世纪 ----------",
                Left = 20,
                Top = 470,
                Width = 390,
                Height = 20,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.DarkGray,
                Font = new Font(Font.FontFamily, 8),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            Controls.Add(copyrightLabel);
        }

        private void LoadAutoStartStatus()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(AutoStartRegKey, false);
                var value = key?.GetValue(AppName);
                _autoStartCheckBox.Checked = value != null;
            }
            catch
            {
                _autoStartCheckBox.Checked = false;
            }
        }

        private void AutoStartCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(AutoStartRegKey, true);
                if (key == null)
                {
                    MessageBox.Show("无法访问注册表", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (_autoStartCheckBox.Checked)
                {
                    // 添加开机自启
                    var exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName;
                    if (!string.IsNullOrEmpty(exePath))
                    {
                        key.SetValue(AppName, $"\"{exePath}\"");
                        MessageBox.Show("已设置开机自动启动", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    // 移除开机自启
                    key.DeleteValue(AppName, false);
                    MessageBox.Show("已取消开机自动启动", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _autoStartCheckBox.Checked = !_autoStartCheckBox.Checked;
            }
        }

        private void BrowseDirButton_Click(object? sender, EventArgs e)
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "选择任务目录",
                SelectedPath = _taskDirInput.Text,
                ShowNewFolderButton = true
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                _taskDirInput.Text = dialog.SelectedPath;
            }
        }

        private void StartButton_Click(object? sender, EventArgs e)
        {
            if (!_service.IsRunning)
            {
                // 验证目录
                var taskDir = _taskDirInput.Text.Trim();
                if (string.IsNullOrEmpty(taskDir))
                {
                    MessageBox.Show("请指定任务目录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                try
                {
                    Directory.CreateDirectory(taskDir);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"无法创建目录: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 启动服务
                _service.Port = (int)_portInput.Value;
                _service.EnableRefresh = _refreshCheckBox.Checked;
                _service.EnablePdf = _pdfCheckBox.Checked;
                _service.TaskDirectory = taskDir;

                // 禁用按钮防止重复点击
                _startButton.Enabled = false;
                _startButton.Text = "启动中...";

                Task.Run(async () =>
                {
                    try
                    {
                        await _service.StartAsync();
                        
                        // 更新UI（需要在UI线程）
                        Invoke(() =>
                        {
                            _startButton.Text = "停止服务";
                            _startButton.Enabled = true;
                            _statusLabel.Text = $"状态: 运行中\nAPI: http://localhost:{_service.Port} (局域网可访问)\n任务目录: {taskDir}";
                            _statusLabel.ForeColor = Color.Green;
                            _portInput.Enabled = false;
                            _taskDirInput.Enabled = false;
                            _browseDirButton.Enabled = false;
                            _refreshCheckBox.Enabled = false;
                            _pdfCheckBox.Enabled = false;
                        });
                    }
                    catch (Exception ex)
                    {
                        Invoke(() =>
                        {
                            MessageBox.Show($"启动失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            _startButton.Text = "启动服务";
                            _startButton.Enabled = true;
                        });
                    }
                });
            }
            else
            {
                // 停止服务
                _startButton.Enabled = false;
                _startButton.Text = "停止中...";

                Task.Run(async () =>
                {
                    try
                    {
                        await _service.StopAsync();
                        
                        // 更新UI（需要在UI线程）
                        Invoke(() =>
                        {
                            _startButton.Text = "启动服务";
                            _startButton.Enabled = true;
                            _statusLabel.Text = "状态: 已停止";
                            _statusLabel.ForeColor = Color.Gray;
                            _portInput.Enabled = true;
                            _taskDirInput.Enabled = true;
                            _browseDirButton.Enabled = true;
                            _refreshCheckBox.Enabled = true;
                            _pdfCheckBox.Enabled = true;
                        });
                    }
                    catch (Exception ex)
                    {
                        Invoke(() =>
                        {
                            MessageBox.Show($"停止失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            _startButton.Enabled = true;
                        });
                    }
                });
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_service.IsRunning)
            {
                Task.Run(async () => await _service.StopAsync()).Wait(3000);
            }
            base.OnFormClosing(e);
        }
    }
}
