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
        private CheckBox _autoDeleteUploadsCheckBox = null!;
        private CheckBox _autoDeleteOutputsCheckBox = null!;
        private NumericUpDown _deleteAfterDaysInput = null!;
        private TextBox _taskDirInput = null!;
        private TextBox _uploadDirInput = null!;
        private TextBox _outputDirInput = null!;
        private Button _browseDirButton = null!;
        private Button _browseUploadDirButton = null!;
        private Button _browseOutputDirButton = null!;
        private Button _startButton = null!;
        private Button _copyUrlButton = null!;
        private Button _copyDocsButton = null!;
        private Button _copyTokenButton = null!;
        private Button _regenerateTokenButton = null!;
        private Label _statusLabel = null!;
        private Label _apiLabel = null!;
        private Label _docsLabel = null!;
        private Label _tokenLabel = null!;
        private TextBox _apiUrlTextBox = null!;
        private TextBox _docsUrlTextBox = null!;
        private TextBox _tokenTextBox = null!;
        private TextBox _logTextBox = null!;
        private NotifyIcon _notifyIcon = null!;
        private string _apiToken = string.Empty;

        private const string AutoStartRegKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string AppName = "WordApiService";

        public MainForm()
        {
            _service = new WordService();
            _service.OnLog += OnServiceLog;
            
            // 生成初始 Token
            _apiToken = GenerateToken();
            
            InitializeUI();  // 这里会加载窗口图标
            InitializeTrayIcon();  // 然后使用窗口图标初始化托盘
            LoadAutoStartStatus();
            
            // 调试：显示嵌入的资源名称
            LogEmbeddedResources();
            
            // 窗口加载完成后自动启动服务
            this.Load += MainForm_Load;
        }

        private async void MainForm_Load(object? sender, EventArgs e)
        {
            // 延迟一小段时间，确保界面完全加载
            await Task.Delay(500);
            
            // 自动启动服务
            OnServiceLog("自动启动服务...");
            StartButton_Click(null, EventArgs.Empty);
            
            // 启动后自动最小化到托盘
            await Task.Delay(1000);
            this.WindowState = FormWindowState.Minimized;
            this.Hide();
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

        private string GenerateToken()
        {
            // 生成 32 字符的随机 Token
            return Guid.NewGuid().ToString("N");
        }

        private void InitializeTrayIcon()
        {
            _notifyIcon = new NotifyIcon();
            
            // 设置托盘图标，使用与窗口相同的图标
            try
            {
                if (this.Icon != null)
                {
                    _notifyIcon.Icon = this.Icon;
                }
                else
                {
                    // 如果窗口图标未加载，尝试从嵌入资源加载
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
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
                            _notifyIcon.Icon = new Icon(stream);
                            break;
                        }
                    }
                    
                    // 如果还是没有，使用系统默认图标
                    if (_notifyIcon.Icon == null)
                    {
                        _notifyIcon.Icon = SystemIcons.Application;
                    }
                }
            }
            catch
            {
                _notifyIcon.Icon = SystemIcons.Application;
            }
            
            _notifyIcon.Text = "Word API 服务";
            _notifyIcon.Visible = true;
            
            // 双击托盘图标显示窗口
            _notifyIcon.DoubleClick += (s, e) =>
            {
                this.Show();
                this.WindowState = FormWindowState.Normal;
                this.Activate();
            };
            
            // 右键菜单
            var contextMenu = new ContextMenuStrip();
            
            var showMenuItem = new ToolStripMenuItem("显示主窗口");
            showMenuItem.Click += (s, e) =>
            {
                this.Show();
                this.WindowState = FormWindowState.Normal;
                this.Activate();
            };
            contextMenu.Items.Add(showMenuItem);
            
            contextMenu.Items.Add(new ToolStripSeparator());
            
            var exitMenuItem = new ToolStripMenuItem("退出");
            exitMenuItem.Click += (s, e) =>
            {
                _notifyIcon.Visible = false;
                Application.Exit();
            };
            contextMenu.Items.Add(exitMenuItem);
            
            _notifyIcon.ContextMenuStrip = contextMenu;
        }

        private void InitializeUI()
        {
            Text = "Word API 服务";
            Width = 500;
            Height = 750;
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

            // Token 标签
            _tokenLabel = new Label
            {
                Text = "API Token:",
                Left = 20,
                Top = 20,
                Width = 80
            };
            Controls.Add(_tokenLabel);

            // Token 文本框
            _tokenTextBox = new TextBox
            {
                Left = 110,
                Top = 20,
                Width = 200,
                Height = 20,
                ReadOnly = true,
                BackColor = Color.White,
                Text = _apiToken
            };
            Controls.Add(_tokenTextBox);

            // 复制 Token 按钮
            _copyTokenButton = new Button
            {
                Text = "复制",
                Left = 320,
                Top = 18,
                Width = 50,
                Height = 24,
                Font = new Font(Font.FontFamily, 8),
                TextAlign = ContentAlignment.MiddleCenter
            };
            _copyTokenButton.Click += CopyTokenButton_Click;
            Controls.Add(_copyTokenButton);

            // 重新生成 Token 按钮
            _regenerateTokenButton = new Button
            {
                Text = "重新生成",
                Left = 380,
                Top = 18,
                Width = 80,
                Height = 24,
                Font = new Font(Font.FontFamily, 8),
                TextAlign = ContentAlignment.MiddleCenter
            };
            _regenerateTokenButton.Click += RegenerateTokenButton_Click;
            Controls.Add(_regenerateTokenButton);

            // 端口配置
            var portLabel = new Label
            {
                Text = "服务端口:",
                Left = 20,
                Top = 55,
                Width = 80
            };
            Controls.Add(portLabel);

            _portInput = new NumericUpDown
            {
                Left = 110,
                Top = 55,
                Width = 290,
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
                Top = 90,
                Width = 80
            };
            Controls.Add(taskDirLabel);

            _taskDirInput = new TextBox
            {
                Left = 110,
                Top = 90,
                Width = 290,
                Text = Path.Combine(AppContext.BaseDirectory, "Tasks")
            };
            Controls.Add(_taskDirInput);

            _browseDirButton = new Button
            {
                Text = "浏览",
                Left = 410,
                Top = 88,
                Width = 50,
                Height = 25
            };
            _browseDirButton.Click += BrowseDirButton_Click;
            Controls.Add(_browseDirButton);

            // 上传目录配置
            var uploadDirLabel = new Label
            {
                Text = "上传目录:",
                Left = 20,
                Top = 125,
                Width = 80
            };
            Controls.Add(uploadDirLabel);

            _uploadDirInput = new TextBox
            {
                Left = 110,
                Top = 125,
                Width = 290,
                Text = Path.Combine(AppContext.BaseDirectory, "Tasks", "uploads")
            };
            Controls.Add(_uploadDirInput);

            _browseUploadDirButton = new Button
            {
                Text = "浏览",
                Left = 410,
                Top = 123,
                Width = 50,
                Height = 25
            };
            _browseUploadDirButton.Click += BrowseUploadDirButton_Click;
            Controls.Add(_browseUploadDirButton);

            // 输出目录配置
            var outputDirLabel = new Label
            {
                Text = "输出目录:",
                Left = 20,
                Top = 160,
                Width = 80
            };
            Controls.Add(outputDirLabel);

            _outputDirInput = new TextBox
            {
                Left = 110,
                Top = 160,
                Width = 290,
                Text = Path.Combine(AppContext.BaseDirectory, "Tasks", "outputs")
            };
            Controls.Add(_outputDirInput);

            _browseOutputDirButton = new Button
            {
                Text = "浏览",
                Left = 410,
                Top = 158,
                Width = 50,
                Height = 25
            };
            _browseOutputDirButton.Click += BrowseOutputDirButton_Click;
            Controls.Add(_browseOutputDirButton);

            // 刷新目录开关
            _refreshCheckBox = new CheckBox
            {
                Text = "启用目录刷新（更新域和目录）",
                Left = 20,
                Top = 200,
                Width = 440,
                Checked = true
            };
            Controls.Add(_refreshCheckBox);

            // 转PDF开关
            _pdfCheckBox = new CheckBox
            {
                Text = "启用 PDF 转换",
                Left = 20,
                Top = 230,
                Width = 440,
                Checked = true
            };
            Controls.Add(_pdfCheckBox);

            // 自动删除上传文件
            _autoDeleteUploadsCheckBox = new CheckBox
            {
                Text = "自动删除上传文件",
                Left = 20,
                Top = 260,
                Width = 200,
                Checked = false
            };
            _autoDeleteUploadsCheckBox.CheckedChanged += AutoDeleteCheckBox_CheckedChanged;
            Controls.Add(_autoDeleteUploadsCheckBox);

            // 自动删除输出文件
            _autoDeleteOutputsCheckBox = new CheckBox
            {
                Text = "自动删除输出文件",
                Left = 240,
                Top = 260,
                Width = 200,
                Checked = false
            };
            _autoDeleteOutputsCheckBox.CheckedChanged += AutoDeleteCheckBox_CheckedChanged;
            Controls.Add(_autoDeleteOutputsCheckBox);

            // 保留天数
            var deleteAfterDaysLabel = new Label
            {
                Text = "保留天数:",
                Left = 20,
                Top = 295,
                Width = 80
            };
            Controls.Add(deleteAfterDaysLabel);

            _deleteAfterDaysInput = new NumericUpDown
            {
                Left = 110,
                Top = 293,
                Width = 80,
                Minimum = 1,
                Maximum = 365,
                Value = 7,
                Enabled = false
            };
            Controls.Add(_deleteAfterDaysInput);

            var daysHintLabel = new Label
            {
                Text = "天（启用自动删除后生效）",
                Left = 200,
                Top = 295,
                Width = 260,
                ForeColor = Color.Gray
            };
            Controls.Add(daysHintLabel);

            // 开机自启开关
            _autoStartCheckBox = new CheckBox
            {
                Text = "开机自动启动",
                Left = 20,
                Top = 330,
                Width = 440,
                Checked = false
            };
            _autoStartCheckBox.CheckedChanged += AutoStartCheckBox_CheckedChanged;
            Controls.Add(_autoStartCheckBox);

            // 启动按钮
            _startButton = new Button
            {
                Text = "启动服务",
                Left = 20,
                Top = 370,
                Width = 440,
                Height = 40
            };
            _startButton.Click += StartButton_Click;
            Controls.Add(_startButton);

            // 状态标签
            _statusLabel = new Label
            {
                Text = "状态: 未启动",
                Left = 20,
                Top = 420,
                Width = 440,
                Height = 20,
                ForeColor = Color.Gray
            };
            Controls.Add(_statusLabel);

            // API 标签
            _apiLabel = new Label
            {
                Text = "API:",
                Left = 20,
                Top = 445,
                Width = 40,
                Height = 20
            };
            Controls.Add(_apiLabel);

            // API URL 文本框
            _apiUrlTextBox = new TextBox
            {
                Left = 60,
                Top = 445,
                Width = 300,
                Height = 20,
                ReadOnly = true,
                BackColor = Color.White,
                Text = ""
            };
            Controls.Add(_apiUrlTextBox);

            // 复制 API 地址按钮
            _copyUrlButton = new Button
            {
                Text = "复制",
                Left = 370,
                Top = 443,
                Width = 60,
                Height = 24,
                Enabled = false,
                Font = new Font(Font.FontFamily, 8),
                TextAlign = ContentAlignment.MiddleCenter
            };
            _copyUrlButton.Click += CopyUrlButton_Click;
            Controls.Add(_copyUrlButton);

            // 文档标签
            _docsLabel = new Label
            {
                Text = "文档:",
                Left = 20,
                Top = 473,
                Width = 40,
                Height = 20
            };
            Controls.Add(_docsLabel);

            // 文档 URL 文本框
            _docsUrlTextBox = new TextBox
            {
                Left = 60,
                Top = 473,
                Width = 300,
                Height = 20,
                ReadOnly = true,
                BackColor = Color.White,
                Text = ""
            };
            Controls.Add(_docsUrlTextBox);

            // 复制文档地址按钮
            _copyDocsButton = new Button
            {
                Text = "复制",
                Left = 370,
                Top = 471,
                Width = 60,
                Height = 24,
                Enabled = false,
                Font = new Font(Font.FontFamily, 8),
                TextAlign = ContentAlignment.MiddleCenter
            };
            _copyDocsButton.Click += CopyDocsButton_Click;
            Controls.Add(_copyDocsButton);

            // 日志显示区域
            var logLabel = new Label
            {
                Text = "运行日志:",
                Left = 20,
                Top = 505,
                Width = 100,
                Height = 20
            };
            Controls.Add(logLabel);

            _logTextBox = new TextBox
            {
                Left = 20,
                Top = 530,
                Width = 440,
                Height = 120,
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
                Top = 660,
                Width = 440,
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
                    OnServiceLog("错误: 无法访问注册表");
                    _autoStartCheckBox.Checked = !_autoStartCheckBox.Checked;
                    return;
                }

                if (_autoStartCheckBox.Checked)
                {
                    // 添加开机自启
                    var exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName;
                    if (!string.IsNullOrEmpty(exePath))
                    {
                        key.SetValue(AppName, $"\"{exePath}\"");
                        OnServiceLog("已设置开机自动启动");
                    }
                }
                else
                {
                    // 移除开机自启
                    key.DeleteValue(AppName, false);
                    OnServiceLog("已取消开机自动启动");
                }
            }
            catch (Exception ex)
            {
                OnServiceLog($"设置开机自启失败: {ex.Message}");
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

        private void BrowseUploadDirButton_Click(object? sender, EventArgs e)
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "选择上传目录",
                SelectedPath = _uploadDirInput.Text,
                ShowNewFolderButton = true
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                _uploadDirInput.Text = dialog.SelectedPath;
            }
        }

        private void BrowseOutputDirButton_Click(object? sender, EventArgs e)
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "选择输出目录",
                SelectedPath = _outputDirInput.Text,
                ShowNewFolderButton = true
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                _outputDirInput.Text = dialog.SelectedPath;
            }
        }

        private void AutoDeleteCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            // 当任一自动删除选项被选中时，启用保留天数输入
            _deleteAfterDaysInput.Enabled = _autoDeleteUploadsCheckBox.Checked || _autoDeleteOutputsCheckBox.Checked;
        }

        private string GetLocalIPAddress()
        {
            try
            {
                var networkInterfaces = System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces();
                
                // 优先选择有默认网关的活动网络接口
                foreach (var ni in networkInterfaces)
                {
                    // 跳过非活动、回环、隧道和虚拟网卡
                    if (ni.OperationalStatus != System.Net.NetworkInformation.OperationalStatus.Up)
                        continue;
                    
                    if (ni.NetworkInterfaceType == System.Net.NetworkInformation.NetworkInterfaceType.Loopback)
                        continue;
                    
                    if (ni.NetworkInterfaceType == System.Net.NetworkInformation.NetworkInterfaceType.Tunnel)
                        continue;
                    
                    // 跳过 VMware 虚拟网卡
                    if (ni.Name.Contains("VMware", StringComparison.OrdinalIgnoreCase))
                        continue;
                    
                    var ipProps = ni.GetIPProperties();
                    
                    // 必须有默认网关（说明是真正的网络连接）
                    if (ipProps.GatewayAddresses.Count == 0)
                        continue;
                    
                    // 获取 IPv4 地址
                    foreach (var ip in ipProps.UnicastAddresses)
                    {
                        if (ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                        {
                            // 排除 169.254.x.x（APIPA 地址）
                            var ipBytes = ip.Address.GetAddressBytes();
                            if (ipBytes[0] == 169 && ipBytes[1] == 254)
                                continue;
                            
                            return ip.Address.ToString();
                        }
                    }
                }
                
                // 如果没有找到有网关的接口，返回第一个非虚拟的 IPv4 地址
                var host = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName());
                foreach (var ip in host.AddressList)
                {
                    if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    {
                        var ipStr = ip.ToString();
                        // 排除虚拟网卡的常见 IP 段
                        if (!ipStr.StartsWith("192.168.206.") && 
                            !ipStr.StartsWith("192.168.153.") &&
                            !ipStr.StartsWith("169.254."))
                        {
                            return ipStr;
                        }
                    }
                }
            }
            catch
            {
                // 忽略错误
            }
            return "127.0.0.1";
        }

        private void CopyUrlButton_Click(object? sender, EventArgs e)
        {
            try
            {
                var localIp = GetLocalIPAddress();
                var url = $"http://{localIp}:{_service.Port}/wordapi";
                Clipboard.SetText(url);
                OnServiceLog($"已复制 API 地址到剪贴板: {url}");
                MessageBox.Show($"已复制 API 地址到剪贴板:\n{url}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"复制失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CopyDocsButton_Click(object? sender, EventArgs e)
        {
            try
            {
                var localIp = GetLocalIPAddress();
                var url = $"http://{localIp}:{_service.Port}/docs";
                Clipboard.SetText(url);
                OnServiceLog($"已复制文档地址到剪贴板: {url}");
                MessageBox.Show($"已复制文档地址到剪贴板:\n{url}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"复制失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CopyTokenButton_Click(object? sender, EventArgs e)
        {
            try
            {
                Clipboard.SetText(_apiToken);
                OnServiceLog($"已复制 Token 到剪贴板");
                MessageBox.Show($"已复制 Token 到剪贴板:\n{_apiToken}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"复制失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RegenerateTokenButton_Click(object? sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "确定要重新生成 Token 吗？\n旧的 Token 将失效。",
                "确认",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (result == DialogResult.Yes)
            {
                _apiToken = GenerateToken();
                _tokenTextBox.Text = _apiToken;
                OnServiceLog($"已重新生成 Token: {_apiToken}");
                MessageBox.Show("Token 已重新生成", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                _service.ApiToken = _apiToken;
                _service.EnableRefresh = _refreshCheckBox.Checked;
                _service.EnablePdf = _pdfCheckBox.Checked;
                _service.TaskDirectory = taskDir;
                _service.UploadDirectory = _uploadDirInput.Text.Trim();
                _service.OutputDirectory = _outputDirInput.Text.Trim();
                _service.AutoDeleteUploads = _autoDeleteUploadsCheckBox.Checked;
                _service.AutoDeleteOutputs = _autoDeleteOutputsCheckBox.Checked;
                _service.DeleteAfterDays = (int)_deleteAfterDaysInput.Value;

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
                            var localIp = GetLocalIPAddress();
                            _startButton.Text = "停止服务";
                            _startButton.Enabled = true;
                            _statusLabel.Text = "状态: 运行中";
                            _statusLabel.ForeColor = Color.Green;
                            _apiUrlTextBox.Text = $"http://{localIp}:{_service.Port}/wordapi";
                            _docsUrlTextBox.Text = $"http://{localIp}:{_service.Port}/docs";
                            _copyUrlButton.Enabled = true;
                            _copyDocsButton.Enabled = true;
                            _portInput.Enabled = false;
                            _taskDirInput.Enabled = false;
                            _uploadDirInput.Enabled = false;
                            _outputDirInput.Enabled = false;
                            _browseDirButton.Enabled = false;
                            _browseUploadDirButton.Enabled = false;
                            _browseOutputDirButton.Enabled = false;
                            _refreshCheckBox.Enabled = false;
                            _pdfCheckBox.Enabled = false;
                            _autoDeleteUploadsCheckBox.Enabled = false;
                            _autoDeleteOutputsCheckBox.Enabled = false;
                            _deleteAfterDaysInput.Enabled = false;
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
                            _apiUrlTextBox.Text = "";
                            _docsUrlTextBox.Text = "";
                            _copyUrlButton.Enabled = false;
                            _copyDocsButton.Enabled = false;
                            _portInput.Enabled = true;
                            _taskDirInput.Enabled = true;
                            _uploadDirInput.Enabled = true;
                            _outputDirInput.Enabled = true;
                            _browseDirButton.Enabled = true;
                            _browseUploadDirButton.Enabled = true;
                            _browseOutputDirButton.Enabled = true;
                            _refreshCheckBox.Enabled = true;
                            _pdfCheckBox.Enabled = true;
                            _autoDeleteUploadsCheckBox.Enabled = true;
                            _autoDeleteOutputsCheckBox.Enabled = true;
                            _deleteAfterDaysInput.Enabled = _autoDeleteUploadsCheckBox.Checked || _autoDeleteOutputsCheckBox.Checked;
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
            // 如果是用户点击关闭按钮，最小化到托盘而不是退出
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
                _notifyIcon.ShowBalloonTip(2000, "Word API 服务", "程序已最小化到系统托盘", ToolTipIcon.Info);
                return;
            }
            
            // 如果是真正退出（从托盘菜单退出），停止服务
            if (_service.IsRunning)
            {
                Task.Run(async () => await _service.StopAsync()).Wait(3000);
            }
            
            _notifyIcon.Visible = false;
            _notifyIcon.Dispose();
            
            base.OnFormClosing(e);
        }
    }
}
