using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;

namespace 多窗口后台模拟点击器
{
    public partial class MainForm : Form
    {
        #region 窗口API声明
        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(Point point);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr ChildWindowFromPointEx(IntPtr hWndParent, Point pt, uint uFlags);

        [DllImport("user32.dll")]
        private static extern bool ScreenToClient(IntPtr hWnd, ref Point lpPoint);

        [DllImport("user32.dll")]
        private static extern bool GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("kernel32.dll")]
        private static extern uint GetLastError();

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        // 添加缺失的POINT结构定义
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        private const uint WM_LBUTTONDOWN = 0x0201;
        private const uint WM_LBUTTONUP = 0x0202;
        private const uint WM_RBUTTONDOWN = 0x0204;
        private const uint WM_RBUTTONUP = 0x0205;
        private const uint WM_MOUSEMOVE = 0x0200;

        private const uint CWP_ALL = 0x0000;
        private const uint CWP_SKIPINVISIBLE = 0x0001;
        private const uint CWP_SKIPDISABLED = 0x0002;
        private const uint CWP_SKIPTRANSPARENT = 0x0004;

        // 热键相关错误代码
        private const int ERROR_HOTKEY_ALREADY_REGISTERED = 1409;

        // 鼠标事件常量
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
        private const uint MOUSEEVENTF_ABSOLUTE = 0x8000;
        #endregion

        // 存储所有点击任务
        private List<ClickTask> clickTasks = new List<ClickTask>();
        // 是否运行中
        private bool isRunning = false;
        // 配置文件路径
        private string configFile = "config.ini";
        // 全局热键ID
        private const int HOTKEY_ID = 100;
        private const int HOTKEY_START = 101;
        private const int HOTKEY_STOP = 102;

        // 添加一个类级别的ListView变量，以便从不同方法中访问
        private DataGridView dataGridTasks;

        // 添加类变量
        private TextBox txtInterval;
        private TextBox txtRandomDelay;
        private Dictionary<int, System.Windows.Forms.Timer> taskTimers = new Dictionary<int, System.Windows.Forms.Timer>();
        private Random random = new Random();

        // 在类成员变量区域添加
        private StatusStrip statusStrip;
        private ToolStripStatusLabel statusLabel;
        private System.Windows.Forms.Timer mouseTrackTimer;
        private CheckBox chkTopMost;
        private System.Windows.Forms.Timer topMostTimer;
        private ContextMenuStrip gridContextMenu;

        // 添加字典存储每个任务的下一次点击间隔
        private Dictionary<int, string> nextClickIntervals = new Dictionary<int, string>();

        public MainForm()
        {
            // 初始化日志系统
            Logger.Initialize();
            Logger.Log(Logger.LogLevel.Info, "MainForm初始化开始");
            
            InitializeComponent();
            RegisterHotKey();
            LoadConfig();
            
            Logger.Log(Logger.LogLevel.Info, "MainForm初始化完成");
        }

        #region 窗体初始化
        private void InitializeComponent()
        {
            // 窗体基本设置
            this.Text = "多窗口后台模拟点击器 0.3 By Ducky錡 + DeepSeek R1 @52pojie";
            this.Size = new Size(620, 380);
            this.StartPosition = FormStartPosition.CenterScreen;
            
            // 去掉this前缀，直接调用方法
            CreateDataGridView();
            CreateControlPanels();
            CreateStatusBar();
            
            // 绑定事件
            this.FormClosing += MainForm_FormClosing;
            
            // 注册Delete键
            // this.KeyPreview = true;  // 删除这行
            // this.KeyDown += MainForm_KeyDown;  // 删除这行
        }

        // 拆分CreateUI方法为几个小方法
        private void CreateDataGridView()
        {
            // 创建任务表格视图
            dataGridTasks = new DataGridView();
            dataGridTasks.Dock = DockStyle.Fill;
            dataGridTasks.AllowUserToAddRows = false;
            dataGridTasks.AllowUserToDeleteRows = false;
            dataGridTasks.ReadOnly = true;
            dataGridTasks.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridTasks.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            
            // 禁止排序
            dataGridTasks.AllowUserToOrderColumns = false;
            
            // 允许调整列宽，但禁止调整行高和表头高度
            dataGridTasks.AllowUserToResizeRows = false;
            dataGridTasks.AllowUserToResizeColumns = true; // 明确允许调整列宽
            dataGridTasks.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing; // 禁止调整表头高度
            
            // 设置行高
            dataGridTasks.RowTemplate.Height = 25; // 固定行高为25像素
            
            // 设置DataGridView的样式
            dataGridTasks.RowHeadersVisible = false;
            dataGridTasks.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
            dataGridTasks.DefaultCellStyle.SelectionBackColor = Color.FromArgb(200, 220, 240);
            dataGridTasks.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridTasks.BackgroundColor = Color.White;
            dataGridTasks.BorderStyle = BorderStyle.Fixed3D;
            dataGridTasks.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dataGridTasks.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dataGridTasks.EnableHeadersVisualStyles = false;
            dataGridTasks.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(220, 220, 220);
            dataGridTasks.ColumnHeadersHeight = 25;
            
            // 添加列并设置宽度
            var colId = new DataGridViewTextBoxColumn();
            colId.Name = "Id";
            colId.HeaderText = "No";
            colId.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colId.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridTasks.Columns.Add(colId);
            
            var colProcessName = new DataGridViewTextBoxColumn();
            colProcessName.Name = "ProcessName";
            colProcessName.HeaderText = "进程";
            //colProcessName.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; // 改为None以便固定宽度
            colProcessName.Width = 8; // 固定宽度为80
            colProcessName.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridTasks.Columns.Add(colProcessName);
            
            var colProcessId = new DataGridViewTextBoxColumn();
            colProcessId.Name = "ProcessId";
            colProcessId.HeaderText = "PID";
            colProcessId.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colProcessId.SortMode = DataGridViewColumnSortMode.NotSortable;
            colProcessId.Visible = false;
            dataGridTasks.Columns.Add(colProcessId);
            
            var colControlText = new DataGridViewTextBoxColumn();
            colControlText.Name = "ControlText";
            colControlText.HeaderText = "控件文本";
            colControlText.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colControlText.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridTasks.Columns.Add(colControlText);
            
            var colControlClass = new DataGridViewTextBoxColumn();
            colControlClass.Name = "ControlClass";
            colControlClass.HeaderText = "控件类型";
            colControlClass.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colControlClass.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridTasks.Columns.Add(colControlClass);
            
            var colControlId = new DataGridViewTextBoxColumn();
            colControlId.Name = "ControlId";
            colControlId.HeaderText = "控件ID";
            colControlId.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colControlId.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridTasks.Columns.Add(colControlId);
            
            var colX = new DataGridViewTextBoxColumn();
            colX.Name = "X";
            colX.HeaderText = "X";
            colX.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colX.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridTasks.Columns.Add(colX);
            
            var colY = new DataGridViewTextBoxColumn();
            colY.Name = "Y";
            colY.HeaderText = "Y";
            colY.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colY.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridTasks.Columns.Add(colY);
            
            var colClickMode = new DataGridViewTextBoxColumn();
            colClickMode.Name = "ClickMode";
            colClickMode.HeaderText = "模式";
            colClickMode.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colClickMode.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridTasks.Columns.Add(colClickMode);
            
            var colClickType = new DataGridViewTextBoxColumn();
            colClickType.Name = "ClickType";
            colClickType.HeaderText = "点击类型";
            colClickType.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colClickType.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridTasks.Columns.Add(colClickType);
            
            var colInterval = new DataGridViewTextBoxColumn();
            colInterval.Name = "Interval";
            colInterval.HeaderText = "点击间隔";
            colInterval.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colInterval.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridTasks.Columns.Add(colInterval);
            
            var colRandomDelay = new DataGridViewTextBoxColumn();
            colRandomDelay.Name = "RandomDelay";
            colRandomDelay.HeaderText = "延迟";
            colRandomDelay.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colRandomDelay.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridTasks.Columns.Add(colRandomDelay);
            
            var colStatus = new DataGridViewTextBoxColumn();
            colStatus.Name = "Status";
            colStatus.HeaderText = "状态";
            colStatus.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colStatus.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridTasks.Columns.Add(colStatus);
            
            // 添加新列显示实际点击间隔
            var colNextInterval = new DataGridViewTextBoxColumn();
            colNextInterval.Name = "NextInterval";
            colNextInterval.HeaderText = "实际间隔";
            colNextInterval.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colNextInterval.SortMode = DataGridViewColumnSortMode.NotSortable;
            colNextInterval.Visible = false;
            dataGridTasks.Columns.Add(colNextInterval);
            
            // 创建上下文菜单
            gridContextMenu = new ContextMenuStrip();
            
            // 添加编辑任务菜单项
            var editItem = new ToolStripMenuItem("编辑任务");
            editItem.Click += EditItem_Click;
            gridContextMenu.Items.Add(editItem);
            
            // 添加删除任务菜单项
            var deleteItem = new ToolStripMenuItem("删除任务");
            deleteItem.Click += DeleteItem_Click;
            gridContextMenu.Items.Add(deleteItem);
            
            // 添加分隔线
            gridContextMenu.Items.Add(new ToolStripSeparator());
            
            // 添加切换状态菜单项
            var toggleItem = new ToolStripMenuItem("切换状态");
            toggleItem.Click += ToggleItemClick;
            gridContextMenu.Items.Add(toggleItem);
            
            // 为DataGridView设置上下文菜单
            dataGridTasks.ContextMenuStrip = gridContextMenu;
            
            // 绑定鼠标右击事件
            dataGridTasks.CellMouseDown += DataGridTasks_CellMouseDown;
            
            // 绑定双击事件处理
            dataGridTasks.CellDoubleClick += (sender, e) => {
                if (e.RowIndex >= 0)
                {
                    // 检查是否点击的是状态列（"状态"列的索引）
                    int statusColIndex = 11; // 修正列索引以考虑隐藏列
                    if (e.ColumnIndex == statusColIndex)
                    {
                        // 双击状态列时切换状态
                        ToggleTaskStatus(e.RowIndex);
                    }
                    else
                    {
                        // 双击其他列时编辑任务
                        EditSelectedTask();
                    }
                }
            };
            
            // 添加到控件集合
            this.Controls.Add(dataGridTasks);
            
            // 为DataGridView添加KeyDown事件处理
            dataGridTasks.KeyDown += DataGridTasks_KeyDown;
            
            // 添加Paint事件处理程序，用于绘制背景提示文字
            dataGridTasks.Paint += DataGridTasks_Paint;
        }

        // 添加Paint事件处理程序
        private void DataGridTasks_Paint(object sender, PaintEventArgs e)
        {
            if (dataGridTasks.Rows.Count == 0)
            {
                // 设置文字样式
                using (var brush = new SolidBrush(Color.FromArgb(120, 120, 120, 120)))
                using (var font = new Font("微软雅黑", 12, FontStyle.Regular))
                {
                    // 计算文字位置，使其居中显示
                    string text = "双击列表可编辑";
                    SizeF textSize = e.Graphics.MeasureString(text, font);
                    float x = (dataGridTasks.Width - textSize.Width) / 2;
                    float y = (dataGridTasks.Height - textSize.Height) / 2;
                    
                    // 绘制文字
                    e.Graphics.DrawString(text, font, brush, x, y);
                }
            }
        }

        private void CreateControlPanels()
        {
            // 修改底部面板布局 - 减小高度
            Panel panelControls = new Panel();
            panelControls.Dock = DockStyle.Bottom;
            panelControls.Height = 80; // 从100减小到80
            panelControls.BackColor = Color.FromArgb(240, 240, 240);
            panelControls.Padding = new Padding(10, 5, 10, 5);
            
            // 创建按钮面板（左侧） - 调整大小和位置
            Panel buttonPanel = new Panel();
            buttonPanel.Location = new Point(10, 5); // 上移位置
            buttonPanel.Size = new Size(200, 70);
            
            // 调整按钮位置和大小
            Button btnToggle = new Button();
            btnToggle.Text = "开始(F9)";
            btnToggle.Size = new Size(90, 28); // 略微加宽按钮
            btnToggle.Location = new Point(0, 5);
            btnToggle.Click += BtnToggle_Click;
            
            // 调整窗口置顶复选框位置
            chkTopMost = new CheckBox();
            chkTopMost.Text = "窗口置顶";
            chkTopMost.AutoSize = true;
            chkTopMost.Location = new Point(8, 46);
            chkTopMost.CheckedChanged += ChkTopMost_CheckedChanged;
            
            Button btnAdd = new Button();
            btnAdd.Text = "添加(F3)";
            btnAdd.Size = new Size(90, 28);
            btnAdd.Location = new Point(100, 5);
            btnAdd.Click += BtnAdd_Click;
            
            Button btnDelete = new Button();
            btnDelete.Text = "删除(Del)";
            btnDelete.Size = new Size(90, 28);
            btnDelete.Location = new Point(100, 38);
            btnDelete.Click += BtnDelete_Click;
            
            // 添加控件到按钮面板
            buttonPanel.Controls.AddRange(new Control[] { btnToggle, chkTopMost, btnAdd, btnDelete });
            
            // 创建设置面板（右侧）
            Panel settingsPanel = new Panel();
            settingsPanel.Dock = DockStyle.Right;
            settingsPanel.Width = 280; // 略微缩小宽度
            settingsPanel.Height = panelControls.Height;
            
            // 调整输入框和标签位置
            Label lblInterval = new Label();
            lblInterval.Text = "点击间隔(毫秒):";
            lblInterval.AutoSize = true;
            lblInterval.Location = new Point(10, 12); // 调整位置
            
            txtInterval = new TextBox();
            txtInterval.Text = "1000";
            txtInterval.Size = new Size(100, 23);
            txtInterval.Location = new Point(120, 8); // 调整位置
            
            Label lblRandomDelay = new Label();
            lblRandomDelay.Text = "随机延迟(毫秒):";
            lblRandomDelay.AutoSize = true;
            lblRandomDelay.Location = new Point(10, 42); // 调整位置
            
            txtRandomDelay = new TextBox();
            txtRandomDelay.Text = "200";
            txtRandomDelay.Size = new Size(100, 23);
            txtRandomDelay.Location = new Point(120, 38); // 调整位置
            
            // 将控件添加到设置面板
            settingsPanel.Controls.AddRange(new Control[] { 
                lblInterval, txtInterval,
                lblRandomDelay, txtRandomDelay
            });
            
            // 将面板添加到底部控制面板
            panelControls.Controls.Add(buttonPanel);
            panelControls.Controls.Add(settingsPanel);
            
            // 将底部控制面板添加到窗体
            this.Controls.Add(panelControls);
            
            // 确保设置面板在最顶层
            settingsPanel.BringToFront();
        }

        private void CreateStatusBar()
        {
            // 创建状态栏
            statusStrip = new StatusStrip();
            statusStrip.SizingGrip = false;
            statusStrip.AutoSize = false;  // 禁用自动大小
            statusStrip.Height = 40;       // 设置足够的高度显示两行
            
            statusLabel = new ToolStripStatusLabel();
            statusLabel.Spring = true;
            statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            statusLabel.Text = "准备就绪";
            statusStrip.Items.Add(statusLabel);
            
            this.Controls.Add(statusStrip);
            
            // 创建鼠标跟踪计时器
            mouseTrackTimer = new System.Windows.Forms.Timer();
            mouseTrackTimer.Interval = 100; // 100毫秒更新一次
            mouseTrackTimer.Tick += MouseTrackTimer_Tick;
            mouseTrackTimer.Start();
        }
        #endregion

        #region 事件处理
        private void BtnToggle_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            
            if (isRunning)
            {
                StopClicking();
                btn.Text = "开始(F9)";
            }
            else
            {
                StartClicking();
                btn.Text = "停止(F9)";
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            // 手动添加任务
            AddNewTask();
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            // 删除选中的任务
            DeleteSelectedTask();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            StopClicking();
            
            // 清理所有计时器
            foreach (var timer in taskTimers.Values)
            {
                timer.Stop();
                timer.Dispose();
            }
            taskTimers.Clear();
            
            SaveConfig();
            UnregisterHotKey();
            
            // 停止鼠标跟踪
            if (mouseTrackTimer != null)
            {
                mouseTrackTimer.Stop();
                mouseTrackTimer.Dispose();
            }
            
            // 清理置顶计时器
            if (topMostTimer != null)
            {
                topMostTimer.Stop();
                topMostTimer.Dispose();
            }
        }

        // 添加DataGridView的KeyDown事件处理程序
        private void DataGridTasks_KeyDown(object sender, KeyEventArgs e)
        {
            // 只处理Delete键
            if (e.KeyCode == Keys.Delete)
            {
                DeleteSelectedTasks();
                e.Handled = true;  // 标记事件已处理
            }
        }
        #endregion

        #region 热键处理
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private void RegisterHotKey()
        {
            // 注册F3键为添加快捷键
            RegisterHotKey(this.Handle, 100, 0, (int)Keys.F3);
            
            // 注册F9键为开始/停止快捷键
            RegisterHotKey(this.Handle, 101, 0, (int)Keys.F9);
        }

        private void UnregisterHotKey()
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            UnregisterHotKey(this.Handle, HOTKEY_START);
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_HOTKEY = 0x0312;
            
            if (m.Msg == WM_HOTKEY)
            {
                if ((int)m.WParam == HOTKEY_ID)
                {
                    // F2热键处理
                    CaptureMousePositionAndWindow();
                }
                else if ((int)m.WParam == HOTKEY_START)
                {
                    // F9热键处理 - 现在是切换行为
                    if (isRunning)
                    {
                        StopClicking();
                        // 找到切换按钮并更新文本
                        foreach (Control control in this.Controls)
                        {
                            if (control is Panel panel)
                            {
                                foreach (Control panelControl in panel.Controls)
                                {
                                    if (panelControl is FlowLayoutPanel flowPanel)
                                    {
                                        foreach (Control c in flowPanel.Controls)
                                        {
                                            if (c is Button btn && (btn.Text == "停止(F9)" || btn.Text == "开始(F9)"))
                                            {
                                                btn.Text = "开始(F9)";
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        StartClicking();
                        // 找到切换按钮并更新文本
                        foreach (Control control in this.Controls)
                        {
                            if (control is Panel panel)
                            {
                                foreach (Control panelControl in panel.Controls)
                                {
                                    if (panelControl is FlowLayoutPanel flowPanel)
                                    {
                                        foreach (Control c in flowPanel.Controls)
                                        {
                                            if (c is Button btn && (btn.Text == "停止(F9)" || btn.Text == "开始(F9)"))
                                            {
                                                btn.Text = "停止(F9)";
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            base.WndProc(ref m);
        }

        private void CaptureMousePositionAndWindow()
        {
            // 获取鼠标当前位置
            Point cursorPos = Cursor.Position;
            
            // 获取鼠标位置下的控件句柄
            IntPtr controlHandle = WindowFromPoint(cursorPos);
            
            if (controlHandle == IntPtr.Zero)
            {
                return;
            }
            
            // 获取父窗口信息
            IntPtr parentWindow = GetParentWindow(controlHandle);
            if (parentWindow == IntPtr.Zero)
            {
                parentWindow = controlHandle;
            }
            
            // 获取进程ID
            uint processId;
            GetWindowThreadProcessId(parentWindow, out processId);
            
            // 获取进程名称
            string processName = "Unknown";
            try
            {
                Process process = Process.GetProcessById((int)processId);
                processName = process.ProcessName;
            }
            catch { }
            
            // 获取控件信息
            StringBuilder controlText = new StringBuilder(256);
            GetWindowText(controlHandle, controlText, controlText.Capacity);
            
            StringBuilder controlClassName = new StringBuilder(256);
            GetClassName(controlHandle, controlClassName, controlClassName.Capacity);
            
            // 获取控件ID
            uint controlProcessId;
            GetWindowThreadProcessId(controlHandle, out controlProcessId);
            
            // 计算控件相对坐标 - 修改这里，使用控件句柄而不是父窗口
            Point controlPoint = cursorPos;
            ScreenToClient(controlHandle, ref controlPoint);
            
            Logger.Log(Logger.LogLevel.Debug, $"添加任务 - 屏幕坐标:({cursorPos.X},{cursorPos.Y}) 控件相对坐标:({controlPoint.X},{controlPoint.Y})");
            
            // 获取文本框中的间隔和随机延迟值
            int interval = 1000;
            int.TryParse(txtInterval.Text, out interval);
            
            int randomDelay = 200;
            int.TryParse(txtRandomDelay.Text, out randomDelay);
            
            // 创建新任务 - 使用控件相对坐标
            ClickTask newTask = new ClickTask
            {
                // 进程信息
                ProcessId = processId,
                ProcessName = processName,
                
                // 控件信息
                ControlHandle = controlHandle,  // 保存控件句柄
                ControlId = controlProcessId,
                ControlText = controlText.ToString(),
                ControlClassName = controlClassName.ToString(),
                
                // 点击信息 - 使用控件相对坐标
                X = controlPoint.X,  // 使用控件相对坐标
                Y = controlPoint.Y,  // 使用控件相对坐标
                ClickType = "左键单击",
                ClickMode = "模式1",
                Interval = interval,
                RandomDelay = randomDelay,
                IsActive = true
            };
            
            // 添加任务并刷新界面
            clickTasks.Add(newTask);
            RefreshTaskListView();
            SaveConfig();
            
            // 更新状态栏，显示控件相对坐标
            statusLabel.Text = $"已添加任务: {processName} | 控件: {controlClassName}({controlText}) | 控件相对坐标: ({controlPoint.X},{controlPoint.Y})";
        }

        private IntPtr FindDeepestChildAtPoint(IntPtr hWndParent, Point screenPoint)
        {
            IntPtr result = hWndParent;
            Point clientPoint = screenPoint;
            
            while (true)
            {
                // 转换为客户区坐标
                ScreenToClient(result, ref clientPoint);
                
                // 查找子控件
                IntPtr child = ChildWindowFromPointEx(
                    result,
                    clientPoint,
                    CWP_SKIPINVISIBLE | CWP_SKIPDISABLED
                );
                
                // 如果没有找到子控件或找到的是自己，则返回当前控件
                if (child == IntPtr.Zero || child == result)
                {
                    return result;
                }
                
                // 继续查找更深层的子控件
                result = child;
                clientPoint = screenPoint;
            }
        }

        private IntPtr GetParentWindow(IntPtr hWnd)
        {
            // 获取窗口所属的顶级窗口
            IntPtr hwndParent = hWnd;
            IntPtr hwndTmp;
            
            while ((hwndTmp = GetParent(hwndParent)) != IntPtr.Zero)
            {
                hwndParent = hwndTmp;
            }
            
            // 如果是同一个窗口，则返回原窗口
            if (hwndParent == hWnd)
                return hWnd;
            
            // 判断父窗口是否为桌面窗口
            StringBuilder className = new StringBuilder(256);
            GetClassName(hwndParent, className, className.Capacity);
            
            // 如果是桌面窗口，返回原窗口
            if (className.ToString() == "Progman" || className.ToString() == "WorkerW")
                return hWnd;
            
            return hwndParent;
        }
        #endregion

        #region 核心功能
        private void StartClicking()
        {
            Logger.Log(Logger.LogLevel.Info, "开始自动点击");
            
            if (isRunning)
            {
                Logger.Log(Logger.LogLevel.Warning, "重复调用StartClicking，程序已经在运行");
                return;
            }
            
            isRunning = true;
            
            // 清理现有计时器
            foreach (var timer in taskTimers.Values)
            {
                timer.Stop();
                timer.Dispose();
            }
            taskTimers.Clear();
            
            nextClickIntervals.Clear();
            
            // 创建新计时器
            for (int i = 0; i < clickTasks.Count; i++)
            {
                var task = clickTasks[i];
                if (task.IsActive)
                {
                    Logger.Log(Logger.LogLevel.Info, $"创建任务计时器 - 索引:{i} 进程:{task.ProcessName} 间隔:{task.Interval}ms");
                    
                    System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                    timer.Interval = task.Interval;
                    timer.Tag = i;
                    timer.Tick += Timer_Tick;
                    taskTimers.Add(i, timer);
                    timer.Start();
                }
                else
                {
                    Logger.Log(Logger.LogLevel.Debug, $"跳过未启用的任务 - 索引:{i} 进程:{task.ProcessName}");
                }
            }
            
            // 更新UI
            UpdateUIStatus();
            RefreshTaskListView();
            
            // 更新按钮文本
            foreach (Control control in this.Controls)
            {
                if (control is Panel panel)
                {
                    foreach (Control panelControl in panel.Controls)
                    {
                        if (panelControl is Button button)
                        {
                            if (button.Text == "开始(F9)")
                            {
                                button.Text = "停止(F9)";
                                Logger.Log(Logger.LogLevel.Debug, "切换按钮文本为 '停止(F9)'");
                            }
                        }
                    }
                }
            }
        }

        private void StopClicking()
        {
            isRunning = false;
            
            // 停止所有计时器
            foreach (var timer in taskTimers.Values)
            {
                timer.Stop();
            }
            
            // 更新UI状态
            UpdateUIStatus();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (!isRunning)
            {
                Logger.Log(Logger.LogLevel.Debug, "Timer_Tick被调用但程序未运行状态");
                return;
            }
            
            System.Windows.Forms.Timer timer = (System.Windows.Forms.Timer)sender;
            int taskIndex = (int)timer.Tag;
            
            Logger.Log(Logger.LogLevel.Debug, $"触发任务计时器 - 索引:{taskIndex}");
            
            if (taskIndex < clickTasks.Count)
            {
                var task = clickTasks[taskIndex];
                
                // 检查窗口是否存在
                bool processExists = false;
                string processName = task.ProcessName ?? "未知进程";
                
                try
                {
                    Process process = Process.GetProcessById((int)task.ProcessId);
                    if (process != null && !process.HasExited)
                    {
                        processExists = true;
                        // 执行点击操作
                        Logger.Log(Logger.LogLevel.Info, $"执行任务点击 - 索引:{taskIndex} 进程:{task.ProcessName} 位置:({task.X},{task.Y})");
                        ExecuteClick(task);
                        
                        // 重要变化：在点击后设置新的随机延迟
                        // 生成0到随机延迟之间的随机数
                        int randomMs = task.RandomDelay > 0 ? random.Next(0, task.RandomDelay + 1) : 0;
                        
                        // 修改计时器间隔为基础间隔加随机延迟
                        timer.Interval = task.Interval + randomMs;
                        Logger.Log(Logger.LogLevel.Debug, $"设置新的计时器间隔 - 任务:{taskIndex} 基础间隔:{task.Interval} 随机延迟:{randomMs} 总间隔:{timer.Interval}");
                        
                        // 保存当前实际间隔信息用于显示
                        nextClickIntervals[taskIndex] = $"{task.Interval}+{randomMs}";
                        
                        // 触发UI刷新
                        this.BeginInvoke(new Action(() => {
                            UpdateNextIntervalDisplay(taskIndex);
                        }));
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log(Logger.LogLevel.Error, $"检查进程存在时出错 - 进程ID:{task.ProcessId} 错误:{ex.Message}");
                    processExists = false;
                }
                
                if (!processExists)
                {
                    Logger.Log(Logger.LogLevel.Warning, $"进程不存在，删除任务 - 进程:{processName} PID:{task.ProcessId}");
                    // 进程不存在，从任务列表中移除
                    timer.Stop();
                    timer.Dispose();
                    taskTimers.Remove(taskIndex);
                    clickTasks.RemoveAt(taskIndex);
                    ReorganizeTimers();
                    RefreshTaskListView();
                    SaveConfig();
                }
            }
            else
            {
                Logger.Log(Logger.LogLevel.Warning, $"无效的任务索引:{taskIndex}，停止计时器");
                // 无效的任务索引，停止计时器
                timer.Stop();
                timer.Dispose();
                taskTimers.Remove(taskIndex);
            }
        }

        private void ExecuteClick(ClickTask task)
        {
            try
            {
                Logger.Log(Logger.LogLevel.Debug, $"开始执行点击 - 进程:{task.ProcessName} PID:{task.ProcessId} 位置:({task.X},{task.Y}) 模式:{task.ClickMode}");
                
                bool success = false;
                
                // 根据不同模式执行不同点击方法
                switch (task.ClickMode)
                {
                    case "模式1":
                        // 方法1: 使用原始句柄直接发送消息
                        if (task.ControlHandle != IntPtr.Zero && IsWindow(task.ControlHandle))
                        {
                            success = SendDirectMessage(task.ControlHandle, task);
                        }
                        break;
                        
                    case "模式2":
                        // 方法2: 尝试找回窗口再点击
                        IntPtr targetWindow = FindTargetWindow(task);
                        if (targetWindow != IntPtr.Zero)
                        {
                            success = SendDirectMessage(targetWindow, task);
                        }
                        break;
                        
                    case "模式3":
                        // 方法3: 使用全局模拟点击
                        success = SendGlobalClick(task);
                        break;
                        
                    case "模式4":
                        // 方法4: 尝试使用SendInput
                        success = TrySendInput(task);
                        break;
                        
                    default:
                        // 默认使用模式1
                        if (task.ControlHandle != IntPtr.Zero && IsWindow(task.ControlHandle))
                        {
                            success = SendDirectMessage(task.ControlHandle, task);
                        }
                        break;
                }
                
                if (success)
                {
                    Logger.Log(Logger.LogLevel.Info, $"点击成功 - {task.ProcessName} ({task.X},{task.Y})");
                    statusLabel.Text = $"点击成功 - {task.ProcessName} ({task.X},{task.Y})";
                }
                else
                {
                    Logger.Log(Logger.LogLevel.Warning, $"点击失败 - {task.ProcessName} ({task.X},{task.Y})");
                    statusLabel.Text = $"点击失败 - {task.ProcessName} ({task.X},{task.Y})";
                }
            }
            catch (Exception ex)
            {
                Logger.Log(Logger.LogLevel.Error, $"点击出错: {ex.Message}");
                statusLabel.Text = $"点击出错: {ex.Message}";
            }
        }

        // 方法1: 直接向控件发送消息
        private bool SendDirectMessage(IntPtr hwnd, ClickTask task)
        {
            try
            {
                Logger.Log(Logger.LogLevel.Debug, $"尝试直接向控件发送消息 - 句柄:0x{hwnd.ToInt64():X8} 相对坐标:({task.X},{task.Y})");
                
                // 构建坐标参数
                int lParam = (task.Y << 16) | (task.X & 0xFFFF);
                
                switch (task.ClickType)
                {
                    case "左键单击":
                        SendMessage(hwnd, WM_LBUTTONDOWN, IntPtr.Zero, (IntPtr)lParam);
                        Thread.Sleep(10);
                        SendMessage(hwnd, WM_LBUTTONUP, IntPtr.Zero, (IntPtr)lParam);
                        break;
                        
                    case "右键单击":
                        SendMessage(hwnd, WM_RBUTTONDOWN, IntPtr.Zero, (IntPtr)lParam);
                        Thread.Sleep(10);
                        SendMessage(hwnd, WM_RBUTTONUP, IntPtr.Zero, (IntPtr)lParam);
                        break;
                        
                    case "左键双击":
                        // 第一次点击
                        SendMessage(hwnd, WM_LBUTTONDOWN, IntPtr.Zero, (IntPtr)lParam);
                        Thread.Sleep(10);
                        SendMessage(hwnd, WM_LBUTTONUP, IntPtr.Zero, (IntPtr)lParam);
                        // 短暂延迟
                        Thread.Sleep(50);
                        // 第二次点击
                        SendMessage(hwnd, WM_LBUTTONDOWN, IntPtr.Zero, (IntPtr)lParam);
                        Thread.Sleep(10);
                        SendMessage(hwnd, WM_LBUTTONUP, IntPtr.Zero, (IntPtr)lParam);
                        break;
                        
                    case "右键双击":
                        // 第一次点击
                        SendMessage(hwnd, WM_RBUTTONDOWN, IntPtr.Zero, (IntPtr)lParam);
                        Thread.Sleep(10);
                        SendMessage(hwnd, WM_RBUTTONUP, IntPtr.Zero, (IntPtr)lParam);
                        // 短暂延迟
                        Thread.Sleep(50);
                        // 第二次点击
                        SendMessage(hwnd, WM_RBUTTONDOWN, IntPtr.Zero, (IntPtr)lParam);
                        Thread.Sleep(10);
                        SendMessage(hwnd, WM_RBUTTONUP, IntPtr.Zero, (IntPtr)lParam);
                        break;
                        
                    case "左键双击(2秒)":
                        // 第一次点击
                        SendMessage(hwnd, WM_LBUTTONDOWN, IntPtr.Zero, (IntPtr)lParam);
                        Thread.Sleep(10);
                        SendMessage(hwnd, WM_LBUTTONUP, IntPtr.Zero, (IntPtr)lParam);
                        // 2秒延迟
                        Thread.Sleep(2000);
                        // 第二次点击
                        SendMessage(hwnd, WM_LBUTTONDOWN, IntPtr.Zero, (IntPtr)lParam);
                        Thread.Sleep(10);
                        SendMessage(hwnd, WM_LBUTTONUP, IntPtr.Zero, (IntPtr)lParam);
                        break;
                        
                    case "右键双击(2秒)":
                        // 第一次点击
                        SendMessage(hwnd, WM_RBUTTONDOWN, IntPtr.Zero, (IntPtr)lParam);
                        Thread.Sleep(10);
                        SendMessage(hwnd, WM_RBUTTONUP, IntPtr.Zero, (IntPtr)lParam);
                        // 2秒延迟
                        Thread.Sleep(2000);
                        // 第二次点击
                        SendMessage(hwnd, WM_RBUTTONDOWN, IntPtr.Zero, (IntPtr)lParam);
                        Thread.Sleep(10);
                        SendMessage(hwnd, WM_RBUTTONUP, IntPtr.Zero, (IntPtr)lParam);
                        break;
                }
                
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log(Logger.LogLevel.Error, $"SendDirectMessage失败: {ex.Message}");
                return false;
            }
        }

        // 方法2: 查找目标窗口
        private IntPtr FindTargetWindow(ClickTask task)
        {
            try
            {
                // 首先尝试通过进程ID找到窗口
                Process[] processes = Process.GetProcessesByName(task.ProcessName);
                foreach (Process proc in processes)
                {
                    if (proc.Id == task.ProcessId)
                    {
                        // 找到了正确的进程
                        Logger.Log(Logger.LogLevel.Debug, $"找到目标进程 - 名称:{task.ProcessName} ID:{task.ProcessId}");
                        return proc.MainWindowHandle;
                    }
                }
                
                // 如果没有找到精确匹配，尝试任何同名进程
                if (processes.Length > 0)
                {
                    Logger.Log(Logger.LogLevel.Debug, $"使用同名进程 - 名称:{task.ProcessName} ID:{processes[0].Id}");
                    return processes[0].MainWindowHandle;
                }
                
                Logger.Log(Logger.LogLevel.Warning, $"未找到目标窗口 - 进程:{task.ProcessName} ID:{task.ProcessId}");
                return IntPtr.Zero;
            }
            catch (Exception ex)
            {
                Logger.Log(Logger.LogLevel.Error, $"FindTargetWindow失败: {ex.Message}");
                return IntPtr.Zero;
            }
        }

        // 方法3: 全局点击
        private bool SendGlobalClick(ClickTask task)
        {
            try
            {
                // 保存当前鼠标位置
                POINT originalPos;
                GetCursorPos(out originalPos);
                
                // 记录日志
                Logger.Log(Logger.LogLevel.Debug, $"执行全局点击 - 原鼠标位置:({originalPos.X},{originalPos.Y}) 目标位置:({task.X},{task.Y})");
                
                // 设置鼠标位置到目标位置
                SetCursorPos(task.X, task.Y);
                Thread.Sleep(10);
                
                // 使用Windows API发送鼠标事件
                if (task.ClickType == "左键单击")
                {
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_ABSOLUTE, task.X, task.Y, 0, UIntPtr.Zero);
                    Thread.Sleep(20);
                    mouse_event(MOUSEEVENTF_LEFTUP | MOUSEEVENTF_ABSOLUTE, task.X, task.Y, 0, UIntPtr.Zero);
                }
                else // 右键单击
                {
                    mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_ABSOLUTE, task.X, task.Y, 0, UIntPtr.Zero);
                    Thread.Sleep(20);
                    mouse_event(MOUSEEVENTF_RIGHTUP | MOUSEEVENTF_ABSOLUTE, task.X, task.Y, 0, UIntPtr.Zero);
                }
                
                // 恢复原来的鼠标位置
                Thread.Sleep(10);
                SetCursorPos(originalPos.X, originalPos.Y);
                
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log(Logger.LogLevel.Error, $"SendGlobalClick失败: {ex.Message}");
                return false;
            }
        }

        // 方法4: 尝试使用SendInput (需要添加相应的API声明)
        private bool TrySendInput(ClickTask task)
        {
            // 我们需要添加SendInput的API声明，暂时使用一个简单版本
            try
            {
                // 这个方法需要额外的API声明和结构定义
                // 简单起见，这里使用方法3作为替代
                POINT originalPos;
                GetCursorPos(out originalPos);
                
                Logger.Log(Logger.LogLevel.Debug, $"尝试使用SendInput模拟点击");
                
                // 强制将当前窗口置为前台窗口
                IntPtr targetWindow = FindTargetWindow(task);
                if (targetWindow != IntPtr.Zero)
                {
                    SetForegroundWindow(targetWindow);
                    Thread.Sleep(50);
                }
                
                SetCursorPos(task.X, task.Y);
                Thread.Sleep(30);
                
                if (task.ClickType == "左键单击")
                {
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_ABSOLUTE, task.X, task.Y, 0, UIntPtr.Zero);
                    Thread.Sleep(30);
                    mouse_event(MOUSEEVENTF_LEFTUP | MOUSEEVENTF_ABSOLUTE, task.X, task.Y, 0, UIntPtr.Zero);
                }
                else // 右键单击
                {
                    mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_ABSOLUTE, task.X, task.Y, 0, UIntPtr.Zero);
                    Thread.Sleep(30);
                    mouse_event(MOUSEEVENTF_RIGHTUP | MOUSEEVENTF_ABSOLUTE, task.X, task.Y, 0, UIntPtr.Zero);
                }
                
                // 恢复原来的鼠标位置
                Thread.Sleep(30);
                SetCursorPos(originalPos.X, originalPos.Y);
                
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log(Logger.LogLevel.Error, $"TrySendInput失败: {ex.Message}");
                return false;
            }
        }

        // 添加一个重新组织计时器的方法
        private void ReorganizeTimers()
        {
            // 停止并清理所有现有计时器
            foreach (var timer in taskTimers.Values)
            {
                timer.Stop();
                timer.Dispose();
            }
            taskTimers.Clear();
            
            // 如果还在运行，为所有任务创建新计时器
            if (isRunning)
            {
                for (int i = 0; i < clickTasks.Count; i++)
                {
                    if (clickTasks[i].IsActive)
                    {
                        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                        timer.Interval = clickTasks[i].Interval;
                        timer.Tag = i;
                        timer.Tick += Timer_Tick;
                        taskTimers.Add(i, timer);
                        timer.Start();
                    }
                }
            }
        }
        #endregion

        #region 辅助方法
        private void AddNewTask()
        {
            try
            {
                // 获取当前鼠标位置
                Point cursorPos = Cursor.Position;
                
                // 获取鼠标位置下的窗口和控件
                IntPtr windowHandle = WindowFromPoint(cursorPos);
                
                // 改进：如果找不到有效窗口，给出提示
                if (windowHandle == IntPtr.Zero)
                {
                    MessageBox.Show("无法获取鼠标下的窗口，请重试", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // 控件句柄就是鼠标下的窗口
                IntPtr controlHandle = windowHandle;
                
                // 寻找顶级窗口
                IntPtr parentWindow = windowHandle;
                while (GetParent(parentWindow) != IntPtr.Zero)
                {
                    parentWindow = GetParent(parentWindow);
                }
                
                if (parentWindow == this.Handle)
                {
                    MessageBox.Show("不能添加本程序窗口", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // 获取控件和窗口信息
                StringBuilder windowTitle = new StringBuilder(256);
                GetWindowText(parentWindow, windowTitle, windowTitle.Capacity);
                
                StringBuilder controlText = new StringBuilder(256);
                GetWindowText(controlHandle, controlText, controlText.Capacity);
                
                StringBuilder controlClassName = new StringBuilder(256);
                GetClassName(controlHandle, controlClassName, controlClassName.Capacity);
                
                uint processId;
                GetWindowThreadProcessId(parentWindow, out processId);
                
                string processName = "未知进程";
                try
                {
                    Process process = Process.GetProcessById((int)processId);
                    processName = process.ProcessName;
                }
                catch { }
                
                // 关键改进：验证并计算控件相对坐标
                Point controlPoint = new Point(cursorPos.X, cursorPos.Y);  // 先复制当前鼠标位置
                bool convertSuccess = ScreenToClient(controlHandle, ref controlPoint);  // 转换为控件相对坐标
                
                if (!convertSuccess)
                {
                    Logger.Log(Logger.LogLevel.Warning, "ScreenToClient失败，尝试使用备用方法计算坐标");
                    // 尝试其他方法获取相对坐标
                    RECT controlRect;
                    if (GetWindowRect(controlHandle, out controlRect))
                    {
                        // 手动计算相对坐标
                        controlPoint.X = cursorPos.X - controlRect.Left;
                        controlPoint.Y = cursorPos.Y - controlRect.Top;
                    }
                }
                
                // 记录坐标转换详细日志
                Logger.Log(Logger.LogLevel.Info, 
                    $"坐标转换 - 屏幕:({cursorPos.X},{cursorPos.Y}) → 控件相对:({controlPoint.X},{controlPoint.Y}), " +
                    $"窗口:0x{parentWindow.ToInt64():X8}, 控件:0x{controlHandle.ToInt64():X8}, 类:{controlClassName}");
                
                // 读取当前设置的间隔和随机延迟
                int interval = 1000;
                int randomDelay = 0;
                
                if (!int.TryParse(txtInterval.Text, out interval)) interval = 1000;
                if (!int.TryParse(txtRandomDelay.Text, out randomDelay)) randomDelay = 0;
                
                // 创建任务
                ClickTask newTask = new ClickTask
                {
                    ProcessId = processId,
                    ProcessName = processName,
                    ControlHandle = controlHandle,  // 保存具体控件句柄
                    ControlText = controlText.ToString(),
                    ControlClassName = controlClassName.ToString(),
                    ControlId = 0,
                    X = controlPoint.X,  // 使用控件相对坐标
                    Y = controlPoint.Y,  // 使用控件相对坐标
                    ClickType = "左键单击",
                    ClickMode = "模式1",
                    Interval = interval,
                    RandomDelay = randomDelay,
                    IsActive = true
                };
                
                clickTasks.Add(newTask);
                RefreshTaskListView();
                SaveConfig();
                
                // 在状态栏显示详细信息
                statusLabel.Text = $"已添加任务: {processName} | 控件: {controlClassName}({controlText}) | 控件相对坐标: ({controlPoint.X},{controlPoint.Y})";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加任务失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logger.Log(Logger.LogLevel.Error, $"添加任务失败: {ex.Message}\n{ex.StackTrace}");
            }
        }

        private void DeleteSelectedTask()
        {
            if (dataGridTasks.SelectedRows.Count > 0)
            {
                int selectedIndex = dataGridTasks.SelectedRows[0].Index;
                if (selectedIndex < clickTasks.Count)
                {
                    // 直接删除任务，无需确认
                    clickTasks.RemoveAt(selectedIndex);
                    
                    // 如果删除的是正在运行的任务，需要停止对应的计时器
                    if (taskTimers.ContainsKey(selectedIndex))
                    {
                        taskTimers[selectedIndex].Stop();
                        taskTimers[selectedIndex].Dispose();
                        taskTimers.Remove(selectedIndex);
                    }
                    
                    // 重新组织计时器索引
                    ReorganizeTimers();
                    
                    // 刷新界面并保存配置
                    RefreshTaskListView();
                    SaveConfig();
                }
            }
            else
            {
                MessageBox.Show("请先选择要删除的项目");
            }
        }

        private void EditSelectedTask()
        {
            if (dataGridTasks.SelectedRows.Count > 0)
            {
                int selectedIndex = dataGridTasks.SelectedRows[0].Index;
                if (selectedIndex < clickTasks.Count)
                {
                    var task = clickTasks[selectedIndex];
                    
                    // 保存当前置顶状态
                    bool originalTopMost = this.TopMost;
                    this.TopMost = false;
                    
                    // 创建编辑窗口 - 调整窗口大小更紧凑
                    Form editForm = new Form();
                    editForm.Text = "编辑任务";
                    editForm.Size = new Size(420, 580); // 调整为更紧凑的尺寸
                    editForm.StartPosition = FormStartPosition.CenterParent;
                    editForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    editForm.MaximizeBox = false;
                    editForm.MinimizeBox = false;
                    editForm.Padding = new Padding(8); // 减小边距
                    
                    // 添加面板来承载所有控件
                    Panel contentPanel = new Panel();
                    contentPanel.Dock = DockStyle.Fill;
                    contentPanel.AutoScroll = true;
                    editForm.Controls.Add(contentPanel);
                    
                    // 进程信息组 - 减小高度和间距
                    GroupBox processGroup = new GroupBox();
                    processGroup.Text = "进程信息";
                    processGroup.Width = 380;
                    processGroup.Height = 85; // 减小高度
                    processGroup.Location = new Point(10, 10);
                    processGroup.Padding = new Padding(8);
                    
                    // 统一标签宽度和对齐方式
                    const int labelWidth = 85;
                    const int controlLeftMargin = 95;
                    
                    Label lblProcessName = new Label { 
                        Text = "进程名称:", 
                        Width = labelWidth,
                        TextAlign = ContentAlignment.MiddleRight,
                        Location = new Point(10, 25)
                    };
                    TextBox txtProcessName = new TextBox { 
                        Text = task.ProcessName,
                        Location = new Point(controlLeftMargin, 25),
                        Width = 260,
                        ReadOnly = true 
                    };
                    
                    Label lblPID = new Label { 
                        Text = "进程ID:", 
                        Width = labelWidth,
                        TextAlign = ContentAlignment.MiddleRight,
                        Location = new Point(10, 50)
                    };
                    TextBox txtPID = new TextBox { 
                        Text = task.ProcessId.ToString(),
                        Location = new Point(controlLeftMargin, 50),
                        Width = 100,
                        ReadOnly = true 
                    };
                    
                    processGroup.Controls.AddRange(new Control[] { lblProcessName, txtProcessName, lblPID, txtPID });
                    
                    // 控件信息组 - 优化布局
                    GroupBox controlGroup = new GroupBox();
                    controlGroup.Text = "控件信息";
                    controlGroup.Width = 380;
                    controlGroup.Height = 120;
                    controlGroup.Location = new Point(10, processGroup.Bottom + 8); // 减小组间距
                    controlGroup.Padding = new Padding(8);
                    
                    Label lblControlText = new Label { 
                        Text = "控件文本:", 
                        Width = labelWidth,
                        TextAlign = ContentAlignment.MiddleRight,
                        Location = new Point(10, 25)
                    };
                    TextBox txtControlText = new TextBox {
                        Text = task.ControlText,
                        Location = new Point(controlLeftMargin, 25),
                        Width = 260,
                        ReadOnly = true
                    };
                    
                    Label lblControlClass = new Label { 
                        Text = "控件类型:", 
                        Width = labelWidth,
                        TextAlign = ContentAlignment.MiddleRight,
                        Location = new Point(10, 50)
                    };
                    TextBox txtControlClass = new TextBox {
                        Text = task.ControlClassName,
                        Location = new Point(controlLeftMargin, 50),
                        Width = 260,
                        ReadOnly = true
                    };
                    
                    Label lblControlId = new Label { 
                        Text = "控件ID:", 
                        Width = labelWidth,
                        TextAlign = ContentAlignment.MiddleRight,
                        Location = new Point(10, 75)
                    };
                    TextBox txtControlId = new TextBox {
                        Text = task.ControlId.ToString(),
                        Location = new Point(controlLeftMargin, 75),
                        Width = 100,
                        ReadOnly = true
                    };
                    
                    controlGroup.Controls.AddRange(new Control[] { 
                        lblControlText, txtControlText,
                        lblControlClass, txtControlClass,
                        lblControlId, txtControlId
                    });
                    
                    // 点击设置组 - 优化布局
                    GroupBox clickGroup = new GroupBox();
                    clickGroup.Text = "点击设置";
                    clickGroup.Width = 380;
                    clickGroup.Height = 220;
                    clickGroup.Location = new Point(10, controlGroup.Bottom + 8);
                    clickGroup.Padding = new Padding(8);
                    
                    // 点击设置组中的控件
                    Label lblX = new Label { 
                        Text = "X坐标:", 
                        Width = labelWidth,
                        TextAlign = ContentAlignment.MiddleRight,
                        Location = new Point(10, 25)
                    };
                    TextBox txtX = new TextBox {
                        Text = task.X.ToString(),
                        Location = new Point(controlLeftMargin, 25),
                        Width = 60
                    };
                    
                    Label lblY = new Label { 
                        Text = "Y坐标:", 
                        Width = labelWidth,
                        TextAlign = ContentAlignment.MiddleRight,
                        Location = new Point(10, 50)
                    };
                    TextBox txtY = new TextBox {
                        Text = task.Y.ToString(),
                        Location = new Point(controlLeftMargin, 50),
                        Width = 60
                    };
                    
                    Label lblClickMode = new Label { 
                        Text = "点击模式:", 
                        Width = labelWidth,
                        TextAlign = ContentAlignment.MiddleRight,
                        Location = new Point(10, 75)
                    };
                    ComboBox cboClickMode = new ComboBox {
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        Location = new Point(controlLeftMargin, 75),
                        Width = 120
                    };
                    cboClickMode.Items.AddRange(new string[] { "模式1", "模式2", "模式3", "模式4" });
                    cboClickMode.SelectedItem = task.ClickMode;
                    
                    Label lblClickType = new Label { 
                        Text = "点击类型:", 
                        Width = labelWidth,
                        TextAlign = ContentAlignment.MiddleRight,
                        Location = new Point(10, 100)
                    };
                    ComboBox cboClickType = new ComboBox {
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        Location = new Point(controlLeftMargin, 100),
                        Width = 120
                    };
                    cboClickType.Items.AddRange(new string[] { 
                        "左键单击", "右键单击",
                        "左键双击", "右键双击",
                        "左键双击(超慢点击,2秒)", "右键双击(超慢点击,2秒)"
                    });
                    cboClickType.SelectedItem = task.ClickType;
                    
                    Label lblInterval = new Label { 
                        Text = "点击间隔:", 
                        Width = labelWidth,
                        TextAlign = ContentAlignment.MiddleRight,
                        Location = new Point(10, 125)
                    };
                    TextBox txtEditInterval = new TextBox {
                        Text = task.Interval.ToString(),
                        Location = new Point(controlLeftMargin, 125),
                        Width = 80
                    };
                    Label lblMs1 = new Label {
                        Text = "毫秒",
                        Location = new Point(controlLeftMargin + 85, 128),
                        AutoSize = true
                    };
                    
                    Label lblDelay = new Label { 
                        Text = "随机延迟:", 
                        Width = labelWidth,
                        TextAlign = ContentAlignment.MiddleRight,
                        Location = new Point(10, 150)
                    };
                    TextBox txtEditDelay = new TextBox {
                        Text = task.RandomDelay.ToString(),
                        Location = new Point(controlLeftMargin, 150),
                        Width = 80
                    };
                    Label lblMs2 = new Label {
                        Text = "毫秒",
                        Location = new Point(controlLeftMargin + 85, 153),
                        AutoSize = true
                    };
                    
                    clickGroup.Controls.AddRange(new Control[] {
                        lblX, txtX,
                        lblY, txtY,
                        lblClickMode, cboClickMode,
                        lblClickType, cboClickType,
                        lblInterval, txtEditInterval, lblMs1,
                        lblDelay, txtEditDelay, lblMs2
                    });
                    
                    // 任务状态复选框
                    CheckBox chkActive = new CheckBox {
                        Text = "启用任务",
                        Checked = task.IsActive,
                        AutoSize = true,
                        Location = new Point(controlLeftMargin, clickGroup.Bottom + 10)
                    };
                    
                    // 按钮布局优化
                    Button btnOK = new Button {
                        Text = "确定",
                        DialogResult = DialogResult.OK,
                        Size = new Size(85, 30),
                        Location = new Point(205, clickGroup.Bottom + 10)
                    };
                    
                    Button btnCancel = new Button {
                        Text = "取消",
                        DialogResult = DialogResult.Cancel,
                        Size = new Size(85, 30),
                        Location = new Point(300, btnOK.Top)
                    };
                    
                    // 添加所有控件到内容面板
                    contentPanel.Controls.AddRange(new Control[] {
                        processGroup, controlGroup, clickGroup,
                        chkActive, btnOK, btnCancel
                    });
                    
                    // 设置默认按钮
                    editForm.AcceptButton = btnOK;
                    editForm.CancelButton = btnCancel;
                    
                    // 显示窗口并处理结果
                    DialogResult result = editForm.ShowDialog(this);
                    
                    // 恢复原来的置顶状态
                    this.TopMost = originalTopMost;
                    
                    // 如果置顶计时器正在运行，确保窗口保持置顶
                    if (originalTopMost && topMostTimer != null && topMostTimer.Enabled)
                    {
                        chkTopMost.Checked = true;
                    }
                    
                    // 处理编辑结果
                    if (result == DialogResult.OK)
                    {
                        // 更新任务数据
                        if (int.TryParse(txtX.Text, out int x)) task.X = x;
                        if (int.TryParse(txtY.Text, out int y)) task.Y = y;
                        if (int.TryParse(txtEditInterval.Text, out int interval)) task.Interval = interval;
                        if (int.TryParse(txtEditDelay.Text, out int delay)) task.RandomDelay = delay;
                        
                        // 更新点击模式
                        if (cboClickMode.SelectedItem != null)
                        {
                            task.ClickMode = cboClickMode.SelectedItem.ToString();
                        }
                        
                        // 更新点击类型
                        if (cboClickType.SelectedItem != null)
                        {
                            task.ClickType = cboClickType.SelectedItem.ToString();
                        }
                        
                        // 更新启用状态
                        task.IsActive = chkActive.Checked;
                        
                        // 刷新列表并保存配置
                        RefreshTaskListView();
                        SaveConfig();
                    }
                }
            }
            else
            {
                MessageBox.Show("请先选择要编辑的项目");
            }
        }

        private DialogResult ShowDialogWithoutTopMost(Form dialogForm)
        {
            // 保存当前置顶状态
            bool originalTopMost = this.TopMost;
            
            // 临时取消置顶状态
            this.TopMost = false;
            
            // 显示对话框
            DialogResult result = dialogForm.ShowDialog(this);
            
            // 恢复原来的置顶状态
            this.TopMost = originalTopMost;
            
            return result;
        }

        private void RefreshTaskListView()
        {
            dataGridTasks.Rows.Clear();
            
            for (int i = 0; i < clickTasks.Count; i++)
            {
                var task = clickTasks[i];
                
                string nextIntervalDisplay = "-";
                if (isRunning && task.IsActive && nextClickIntervals.ContainsKey(i))
                {
                    nextIntervalDisplay = nextClickIntervals[i];
                }
                
                dataGridTasks.Rows.Add(
                    (i + 1).ToString(),
                    task.ProcessName,
                    task.ProcessId.ToString(),
                    task.ControlText,
                    task.ControlClassName,
                    task.ControlId.ToString(),
                    task.X.ToString(),
                    task.Y.ToString(),
                    task.ClickMode,
                    task.ClickType,
                    task.Interval.ToString(),
                    task.RandomDelay.ToString(),
                    task.IsActive ? "启用" : "禁用",
                    nextIntervalDisplay
                );
                
                // 如果程序正在运行，同步更新计时器状态
                if (isRunning)
                {
                    if (task.IsActive)
                    {
                        if (!taskTimers.ContainsKey(i))
                        {
                            // 创建新计时器
                            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                            timer.Interval = task.Interval;
                            timer.Tag = i;
                            timer.Tick += Timer_Tick;
                            taskTimers.Add(i, timer);
                            timer.Start();
                        }
                        else
                        {
                            // 更新计时器间隔
                            taskTimers[i].Interval = task.Interval;
                            taskTimers[i].Start();
                        }
                    }
                    else if (taskTimers.ContainsKey(i))
                    {
                        // 停止非活动任务的计时器
                        taskTimers[i].Stop();
                    }
                }
            }
            
            // 清理无效的计时器
            List<int> keysToRemove = new List<int>();
            foreach (var key in taskTimers.Keys)
            {
                if (key >= clickTasks.Count)
                {
                    taskTimers[key].Stop();
                    taskTimers[key].Dispose();
                    keysToRemove.Add(key);
                }
            }
            
            foreach (var key in keysToRemove)
            {
                taskTimers.Remove(key);
            }
            
            // 数据加载完成后自动调整列宽
            dataGridTasks.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            
            // 设置最小列宽
            foreach (DataGridViewColumn column in dataGridTasks.Columns)
            {
                if (column.Width < 50)
                {
                    column.Width = 50; // 设置最小宽度为50像素
                }
            }
        }

        private void UpdateUIStatus()
        {
            // 根据程序运行状态更新UI
            foreach (Control control in this.Controls)
            {
                if (control is Panel panel)
                {
                    foreach (Control panelControl in panel.Controls)
                    {
                        if (panelControl is Button button)
                        {
                            if (button.Text == "开始")
                            {
                                button.Enabled = !isRunning;
                            }
                            else if (button.Text == "停止")
                            {
                                button.Enabled = isRunning;
                            }
                        }
                    }
                }
            }
            
            // 更新状态显示
            this.Text = isRunning ? "多窗口后台模拟点击器 [运行中]" : "多窗口后台模拟点击器 [已停止]";
        }

        private void DataGridTasks_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // 设置选中行
                dataGridTasks.Rows[e.RowIndex].Selected = true;
            }
        }

        // 添加鼠标右击事件处理
        private void DataGridTasks_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            // 检查是否是右键点击并且点击在有效的行上
            if (e.Button == MouseButtons.Right && e.RowIndex >= 0)
            {
                // 选中当前行
                dataGridTasks.ClearSelection();
                dataGridTasks.Rows[e.RowIndex].Selected = true;
            }
        }

        // 添加菜单项点击事件处理
        private void ToggleItemClick(object sender, EventArgs e)
        {
            if (dataGridTasks.SelectedRows.Count > 0)
            {
                int selectedIndex = dataGridTasks.SelectedRows[0].Index;
                if (selectedIndex < clickTasks.Count)
                {
                    // 切换选中任务的启用状态
                    clickTasks[selectedIndex].IsActive = !clickTasks[selectedIndex].IsActive;
                    
                    // 更新UI和配置
                    RefreshTaskListView();
                    SaveConfig();
                }
            }
        }

        // 添加切换状态的方法
        private void ToggleTaskStatus(int rowIndex)
        {
            if (rowIndex < clickTasks.Count)
            {
                // 切换任务状态
                clickTasks[rowIndex].IsActive = !clickTasks[rowIndex].IsActive;
                
                // 更新UI和配置
                RefreshTaskListView();
                SaveConfig();
            }
        }

        // 添加更新单个单元格的方法
        private void UpdateNextIntervalDisplay(int taskIndex)
        {
            if (taskIndex < dataGridTasks.Rows.Count)
            {
                // 更新"实际间隔"列的显示
                int colIndex = dataGridTasks.Columns.Count - 1; // 最后一列
                if (nextClickIntervals.ContainsKey(taskIndex))
                {
                    dataGridTasks.Rows[taskIndex].Cells[colIndex].Value = nextClickIntervals[taskIndex];
                }
                else
                {
                    dataGridTasks.Rows[taskIndex].Cells[colIndex].Value = "-";
                }
            }
        }
        #endregion

        #region 配置文件处理
        private void SaveConfig()
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(configFile))
                {
                    writer.WriteLine($"[Config]");
                    writer.WriteLine($"TaskCount={clickTasks.Count}");
                    writer.WriteLine($"DefaultInterval={txtInterval.Text}");
                    writer.WriteLine($"DefaultRandomDelay={txtRandomDelay.Text}");
                    
                    for (int i = 0; i < clickTasks.Count; i++)
                    {
                        var task = clickTasks[i];
                        writer.WriteLine($"[Task{i}]");
                        writer.WriteLine($"ProcessName={task.ProcessName}");
                        writer.WriteLine($"ProcessId={task.ProcessId}");
                        writer.WriteLine($"ControlText={task.ControlText}");
                        writer.WriteLine($"ControlClassName={task.ControlClassName}");
                        writer.WriteLine($"ControlId={task.ControlId}");
                        writer.WriteLine($"X={task.X}");
                        writer.WriteLine($"Y={task.Y}");
                        writer.WriteLine($"ClickMode={task.ClickMode}");
                        writer.WriteLine($"ClickType={task.ClickType}");
                        writer.WriteLine($"Interval={task.Interval}");
                        writer.WriteLine($"RandomDelay={task.RandomDelay}");
                        writer.WriteLine($"IsActive={task.IsActive}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存配置文件失败: {ex.Message}");
            }
        }

        private void LoadConfig()
        {
            if (!File.Exists(configFile))
                return;
                
            try
            {
                clickTasks.Clear();
                
                string[] lines = File.ReadAllLines(configFile);
                int taskCount = 0;
                ClickTask currentTask = null;
                
                foreach (string line in lines)
                {
                    if (line.StartsWith("[Config]"))
                    {
                        continue;
                    }
                    else if (line.StartsWith("TaskCount="))
                    {
                        taskCount = int.Parse(line.Substring("TaskCount=".Length));
                    }
                    else if (line.StartsWith("DefaultInterval="))
                    {
                        txtInterval.Text = line.Substring("DefaultInterval=".Length);
                    }
                    else if (line.StartsWith("DefaultRandomDelay="))
                    {
                        txtRandomDelay.Text = line.Substring("DefaultRandomDelay=".Length);
                    }
                    else if (line.StartsWith("[Task"))
                    {
                        if (currentTask != null)
                        {
                            clickTasks.Add(currentTask);
                        }
                        currentTask = new ClickTask();
                    }
                    else if (currentTask != null)
                    {
                        if (line.StartsWith("ProcessName="))
                        {
                            currentTask.ProcessName = line.Substring("ProcessName=".Length);
                        }
                        else if (line.StartsWith("ProcessId="))
                        {
                            currentTask.ProcessId = uint.Parse(line.Substring("ProcessId=".Length));
                        }
                        else if (line.StartsWith("ControlText="))
                        {
                            currentTask.ControlText = line.Substring("ControlText=".Length);
                        }
                        else if (line.StartsWith("ControlClassName="))
                        {
                            currentTask.ControlClassName = line.Substring("ControlClassName=".Length);
                        }
                        else if (line.StartsWith("ControlId="))
                        {
                            currentTask.ControlId = uint.Parse(line.Substring("ControlId=".Length));
                        }
                        else if (line.StartsWith("X="))
                        {
                            currentTask.X = int.Parse(line.Substring("X=".Length));
                        }
                        else if (line.StartsWith("Y="))
                        {
                            currentTask.Y = int.Parse(line.Substring("Y=".Length));
                        }
                        else if (line.StartsWith("ClickMode="))
                        {
                            currentTask.ClickMode = line.Substring("ClickMode=".Length);
                        }
                        else if (line.StartsWith("ClickType="))
                        {
                            currentTask.ClickType = line.Substring("ClickType=".Length);
                        }
                        else if (line.StartsWith("Interval="))
                        {
                            currentTask.Interval = int.Parse(line.Substring("Interval=".Length));
                        }
                        else if (line.StartsWith("RandomDelay="))
                        {
                            currentTask.RandomDelay = int.Parse(line.Substring("RandomDelay=".Length));
                        }
                        else if (line.StartsWith("IsActive="))
                        {
                            currentTask.IsActive = bool.Parse(line.Substring("IsActive=".Length));
                        }
                    }
                }
                
                // 添加最后一个任务
                if (currentTask != null)
                {
                    clickTasks.Add(currentTask);
                }
                
                // 尝试查找窗口句柄
                foreach (var task in clickTasks)
                {
                    task.ControlHandle = FindControlByProcessId(task.ProcessId, task.ControlId);
                }
                
                RefreshTaskListView();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载配置文件失败: {ex.Message}");
            }
        }

        private IntPtr FindControlByProcessId(uint processId, uint controlId)
        {
            try
            {
                Process process = Process.GetProcessById((int)processId);
                if (process.MainWindowHandle != IntPtr.Zero)
                {
                    // 简单情况：返回主窗口句柄
                    return process.MainWindowHandle;
                }
            }
            catch
            {
                // 进程不存在或没有权限访问
            }
            
            return IntPtr.Zero;
        }
        #endregion

        // 添加鼠标跟踪定时器事件处理方法
        private void MouseTrackTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                // 获取鼠标当前位置
                Point cursorPos = Cursor.Position;
                
                // 获取鼠标位置下的控件句柄
                IntPtr controlHandle = WindowFromPoint(cursorPos);
                
                if (controlHandle == IntPtr.Zero)
                {
                    statusLabel.Text = "未能获取鼠标位置下的控件";
                    return;
                }
                
                // 获取父窗口
                IntPtr parentWindow = GetParentWindow(controlHandle);
                if (parentWindow == IntPtr.Zero)
                {
                    parentWindow = controlHandle;
                }
                
                // 获取窗口标题并限制长度
                StringBuilder windowTitle = new StringBuilder(256);
                GetWindowText(parentWindow, windowTitle, windowTitle.Capacity);
                string limitedWindowTitle = windowTitle.ToString();
                if (limitedWindowTitle.Length > 16)
                {
                    limitedWindowTitle = limitedWindowTitle.Substring(0, 13) + "...";
                }
                
                // 获取控件信息
                StringBuilder controlText = new StringBuilder(256);
                GetWindowText(controlHandle, controlText, controlText.Capacity);
                
                StringBuilder controlClassName = new StringBuilder(256);
                GetClassName(controlHandle, controlClassName, controlClassName.Capacity);
                
                // 获取进程信息
                uint processId;
                GetWindowThreadProcessId(parentWindow, out processId);
                
                // 获取进程名称
                string processName = "Unknown";
                try
                {
                    Process process = Process.GetProcessById((int)processId);
                    processName = process.ProcessName.Substring(0, Math.Min(process.ProcessName.Length, 16));
                }
                catch { }
                
                // 获取控件ID
                uint controlProcessId;
                GetWindowThreadProcessId(controlHandle, out controlProcessId);
                
                // 计算窗口相对坐标
                Point windowPoint = cursorPos;
                ScreenToClient(parentWindow, ref windowPoint);
                
                // 计算控件相对坐标
                Point controlPoint = cursorPos;
                ScreenToClient(controlHandle, ref controlPoint);
                
                // 构建第一行信息 - 进程、窗口和控件信息
                string infoLine = string.Format(
                    "进程: {0}({1}) | 窗口: {2}(0x{3:X8}) | 控件: {4}({5},0x{6:X8})",
                    processName,
                    processId,
                    limitedWindowTitle,  // 使用限制长度后的窗口标题
                    parentWindow.ToInt64(),
                    controlText.ToString(),
                    controlClassName.ToString(),
                    controlHandle.ToInt64()
                );
                
                // 构建第二行信息 - 坐标信息
                string coordsLine = string.Format(
                    "屏幕绝对坐标({0},{1}) | 窗口相对坐标({2},{3}) | 控件相对坐标({4},{5})",
                    cursorPos.X,
                    cursorPos.Y,
                    windowPoint.X,
                    windowPoint.Y,
                    controlPoint.X,
                    controlPoint.Y
                );
                
                // 合并两行信息并显示
                statusLabel.Text = infoLine + Environment.NewLine + coordsLine;
            }
            catch (Exception ex)
            {
                statusLabel.Text = "获取鼠标信息错误: " + ex.Message;
            }
        }

        // 添加复选框状态变化事件处理
        private void ChkTopMost_CheckedChanged(object sender, EventArgs e)
        {
            if (chkTopMost.Checked)
            {
                // 立即设置窗口置顶
                this.TopMost = true;
                
                // 创建并启动计时器，每2秒重新设置一次置顶
                if (topMostTimer == null)
                {
                    topMostTimer = new System.Windows.Forms.Timer();
                    topMostTimer.Interval = 2000; // 2000毫秒
                    topMostTimer.Tick += TopMostTimer_Tick;
                }
                topMostTimer.Start();
            }
            else
            {
                // 取消窗口置顶
                this.TopMost = false;
                
                // 停止计时器
                if (topMostTimer != null)
                {
                    topMostTimer.Stop();
                }
            }
        }

        // 添加窗口置顶计时器事件处理
        private void TopMostTimer_Tick(object sender, EventArgs e)
        {
            // 确保窗口保持置顶
            if (chkTopMost.Checked && !this.TopMost)
            {
                this.TopMost = true;
            }
        }

        // 添加编辑任务菜单项的点击事件处理程序
        private void EditItem_Click(object sender, EventArgs e)
        {
            // 调用现有的编辑任务方法
            EditSelectedTask();
        }

        // 添加删除任务菜单项的点击事件处理程序
        private void DeleteItem_Click(object sender, EventArgs e)
        {
            // 调用现有的删除任务方法
            DeleteSelectedTasks();
        }

        // 添加删除任务的方法
        private void DeleteSelectedTasks()
        {
            // 检查是否有选中的行
            if (dataGridTasks.SelectedRows.Count > 0)
            {
                // 确认是否要删除
                DialogResult result = MessageBox.Show(
                    "确定要删除选中的任务吗？",
                    "确认删除",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );
                
                if (result == DialogResult.Yes)
                {
                    // 从后向前删除选中的行，避免索引变化的问题
                    for (int i = dataGridTasks.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        int index = dataGridTasks.SelectedRows[i].Index;
                        if (index < clickTasks.Count)
                        {
                            // 如果任务正在运行，先停止相关的计时器
                            if (taskTimers.ContainsKey(index))
                            {
                                taskTimers[index].Stop();
                                taskTimers[index].Dispose();
                                taskTimers.Remove(index);
                            }
                            
                            // 从任务列表中删除
                            clickTasks.RemoveAt(index);
                            
                            // 从下一次点击间隔字典中删除
                            if (nextClickIntervals.ContainsKey(index))
                            {
                                nextClickIntervals.Remove(index);
                            }
                        }
                    }
                    
                    // 刷新列表显示
                    RefreshTaskListView();
                    
                    // 保存配置
                    SaveConfig();
                    
                    // 更新状态栏
                    statusLabel.Text = "已删除选中的任务";
                    Logger.Log(Logger.LogLevel.Info, "已删除选中的任务");
                }
            }
            else
            {
                MessageBox.Show("请先选择要删除的任务", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    // 点击任务类
    public class ClickTask
    {
        public uint ProcessId { get; set; }            // 替代WindowHandle
        public string ProcessName { get; set; }        // 替代WindowName
        public IntPtr ControlHandle { get; set; }
        public uint ControlId { get; set; }
        public string ControlText { get; set; }
        public string ControlClassName { get; set; }
        public int X { get; set; }
        public int Y { get; set; }
        public string ClickType { get; set; }
        public int Interval { get; set; }
        public int RandomDelay { get; set; }
        public bool IsActive { get; set; }
        public string ClickMode { get; set; } = "模式1"; // 默认使用模式1
    }

    // 添加日志类
    public static class Logger
    {
        private static string logFilePath;
        private static object lockObj = new object();
        private static bool initialized = false;

        // 日志级别
        public enum LogLevel
        {
            Debug,
            Info,
            Warning,
            Error,
            Fatal
        }

        // 初始化日志系统
        public static void Initialize()
        {
            if (initialized) return;

            try
            {
                string logDir = Path.Combine(Application.StartupPath, "Logs");
                
                // 创建日志目录
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }
                
                // 创建日志文件名，包含日期
                string fileName = $"AutoClick_{DateTime.Now:yyyy-MM-dd}.log";
                logFilePath = Path.Combine(logDir, fileName);
                
                // 在日志文件开头记录程序启动信息
                Log(LogLevel.Info, "==================================================");
                Log(LogLevel.Info, $"程序启动 - 版本: 1.0 - 时间: {DateTime.Now}");
                Log(LogLevel.Info, "==================================================");
                
                initialized = true;
            }
            catch (Exception ex)
            {
                // 如果初始化失败，尝试使用备用路径
                logFilePath = Path.Combine(Application.StartupPath, "AutoClick.log");
                Log(LogLevel.Error, $"日志系统初始化失败: {ex.Message}");
            }
        }

        // 记录日志
        public static void Log(LogLevel level, string message)
        {
            try
            {
                lock (lockObj)
                {
                    if (string.IsNullOrEmpty(logFilePath))
                    {
                        Initialize();
                    }
                    
                    // 格式化日志消息
                    string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}";
                    
                    // 写入文件
                    using (StreamWriter writer = new StreamWriter(logFilePath, true))
                    {
                        writer.WriteLine(logMessage);
                    }
                }
            }
            catch
            {
                // 日志记录失败时不做任何处理，避免影响主程序
            }
        }
    }
} 