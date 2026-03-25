using AutoLiquid_GenScript_Single_Handling.EntityJson;
using MahApps.Metro.Controls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace AutoLiquid_GenScript_Single_Handling.Window
{
    /// <summary>
    /// EP管架次级扫码窗口。
    /// 根据耗材 RowCount × ColCount 动态绘制孔位网格，
    /// 有 primerLabel 的孔高亮显示并按顺序逐管扫码；
    /// 所有管通过后自动关闭；点取消则 IsConfirmed=false。
    /// 源盘或靶盘均适用。
    /// </summary>
    public partial class WindowEpRackScan : MetroWindow
    {
        // ── Win32 API ──
        [DllImport("user32.dll")]
        private static extern bool GetKeyboardState(byte[] lpKeyState);
        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int ToUnicode(
            uint wVirtKey, uint wScanCode,
            byte[] lpKeyState,
            [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pwszBuff,
            int cchBuff, uint wFlags);

        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int VK_RETURN = 0x0D;
        private const int VK_BACK = 0x08;
        private static readonly HashSet<int> _ignoredVk = new HashSet<int>
        {
            0x10, 0x11, 0x12, 0x14,
            0x5B, 0x5C,
            0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5,
            0x09, 0x1B
        };

        // ── 孔位数据模型 ──
        /// <summary>单个孔位的扫码信息</summary>
        private class HoleScanItem
        {
            /// <summary>孔位Index（0-based，列优先）</summary>
            public int HoleIndex { get; set; }
            /// <summary>孔位名，如 A1、B1</summary>
            public string HoleName { get; set; }
            /// <summary>期望引物名称（来自Excel A列），空=无需扫码</summary>
            public string ExpectedPrimerLabel { get; set; }
            /// <summary>是否需要扫码（primerLabel非空）</summary>
            public bool NeedScan => !string.IsNullOrEmpty(ExpectedPrimerLabel);
            /// <summary>null=未扫, true=通过, false=不匹配</summary>
            public bool? ScanResult { get; set; } = null;
        }

        // ── 字段 ──
        private readonly int _rowCount;
        private readonly int _colCount;

        /// <summary>key = holeIndex（0-based，列优先）</summary>
        private readonly List<HoleScanItem> _allHoles = new List<HoleScanItem>();

        /// <summary>需要扫码的孔，按顺序排列</summary>
        private readonly Queue<int> _pendingQueue = new Queue<int>(); // 存 _allHoles 的 index

        private int _currentListIndex = -1; // 当前在 _allHoles 中的 index

        // 网格 UI 引用，key = holeIndex
        private readonly Dictionary<int, Border> _holeBorders = new Dictionary<int, Border>();
        private readonly Dictionary<int, TextBlock> _holeStatusTexts = new Dictionary<int, TextBlock>();

        private readonly StringBuilder _inputBuffer = new StringBuilder();
        private bool _hookRegistered = false;

        /// <summary>是否全部确认（供主框判断）</summary>
        public bool IsConfirmed { get; private set; } = false;

        // ── 构造 ──
        /// <param name="consumable">EP管架耗材配置（提供 RowCount / ColCount）</param>
        /// <param name="slotTitle">盘位标题，如"源盘1"或"靶盘1"</param>
        /// <param name="primerLabels">
        ///   按孔Index（0-based，列优先）顺序的引物名称列表。
        ///   空字符串 = 该孔无对应Excel行，不扫码但仍绘制格子。
        /// </param>
        public WindowEpRackScan(Consumable consumable, string slotTitle, List<string> primerLabels)
        {
            InitializeComponent();

            _rowCount = consumable.RowCount;
            _colCount = consumable.ColCount;

            var logoFile = AppDomain.CurrentDomain.BaseDirectory
                           + System.IO.Path.DirectorySeparatorChar + "logo.png";
            if (File.Exists(logoFile))
                this.Icon = new BitmapImage(new Uri(logoFile));

            this.TextBlockTitle.Text = string.Format(
                (string)Application.Current.FindResource("EpRackScanHint"),
                slotTitle);

            BuildHoles(primerLabels);
            BuildGrid();
            InitQueue();
            RefreshProgress();

            this.BtnCancel.Click += (s, e) => { IsConfirmed = false; this.Close(); };
            this.SourceInitialized += OnSourceInitialized;
            this.Closed += OnClosed;
        }

        // ══════════════════════════════════════════════
        // 数据初始化
        // ══════════════════════════════════════════════

        /// <summary>
        /// 根据 RowCount × ColCount 建立全量孔位列表（列优先，与主框一致）。
        /// </summary>
        private void BuildHoles(List<string> primerLabels)
        {
            _allHoles.Clear();
            int total = _rowCount * _colCount;
            for (int holeIdx = 0; holeIdx < total; holeIdx++)
            {
                int rowIdx = holeIdx % _rowCount;
                int colIdx = holeIdx / _rowCount;
                string holeName = (char)('A' + rowIdx) + (colIdx + 1).ToString();

                string primer = (holeIdx < primerLabels.Count) ? primerLabels[holeIdx] : "";

                _allHoles.Add(new HoleScanItem
                {
                    HoleIndex = holeIdx,
                    HoleName = holeName,
                    ExpectedPrimerLabel = primer,
                    ScanResult = null
                });
            }
        }

        // ══════════════════════════════════════════════
        // 网格绘制（与主框 BuildLayoutGrid 风格一致）
        // ══════════════════════════════════════════════

        /// <summary>
        /// 动态绘制 _rowCount 行 × _colCount 列的孔位网格。
        /// 网格按列排列孔位名（A列纵向，1列横向），
        /// 第一行为列号表头，第一列为行字母表头。
        /// 每个孔格显示：孔位名 + 引物名（若有）+ 扫码状态。
        /// </summary>
        private void BuildGrid()
        {
            var grid = this.GridEpRack;
            grid.RowDefinitions.Clear();
            grid.ColumnDefinitions.Clear();
            grid.Children.Clear();
            _holeBorders.Clear();
            _holeStatusTexts.Clear();

            // 表头行 + 数据行
            // Row 0: 列号表头；Row 1~_rowCount: 数据行（A/B/C...）
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 表头
            for (int r = 0; r < _rowCount; r++)
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // 表头列 + 数据列
            // Col 0: 行字母表头；Col 1~_colCount: 数据列（1/2/3...）
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // 行字母表头
            for (int c = 0; c < _colCount; c++)
                grid.ColumnDefinitions.Add(new ColumnDefinition
                { Width = new GridLength(1, GridUnitType.Star) });

            // ── 列号表头（第0行）──
            for (int c = 0; c < _colCount; c++)
            {
                var header = new TextBlock
                {
                    Text = (c + 1).ToString(),
                    FontSize = 10,
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.Gray,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(2, 2, 2, 2)
                };
                Grid.SetRow(header, 0);
                Grid.SetColumn(header, c + 1);
                grid.Children.Add(header);
            }

            // ── 行字母表头（第0列）──
            for (int r = 0; r < _rowCount; r++)
            {
                var header = new TextBlock
                {
                    Text = ((char)('A' + r)).ToString(),
                    FontSize = 10,
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.Gray,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(4, 2, 4, 2)
                };
                Grid.SetRow(header, r + 1);
                Grid.SetColumn(header, 0);
                grid.Children.Add(header);
            }

            // ── 孔位格子（列优先：holeIdx = col * rowCount + row）──
            for (int c = 0; c < _colCount; c++)
            {
                for (int r = 0; r < _rowCount; r++)
                {
                    int holeIdx = c * _rowCount + r;
                    var hole = _allHoles[holeIdx];

                    // 背景色：有引物 = 白色；无引物 = 浅灰
                    var bgColor = hole.NeedScan
                        ? Color.FromRgb(255, 255, 255)
                        : Color.FromRgb(240, 240, 240);

                    var border = new Border
                    {
                        BorderBrush = hole.NeedScan ? Brushes.SteelBlue : Brushes.LightGray,
                        BorderThickness = new Thickness(1),
                        Margin = new Thickness(2),
                        MinWidth = 60,
                        MinHeight = 52,
                        Background = new SolidColorBrush(bgColor)
                    };

                    var panel = new StackPanel
                    {
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Margin = new Thickness(3, 2, 3, 2)
                    };

                    // 孔位名（如 A1）
                    panel.Children.Add(new TextBlock
                    {
                        Text = hole.HoleName,
                        FontSize = 10,
                        FontWeight = FontWeights.Bold,
                        Foreground = Brushes.DarkSlateGray,
                        HorizontalAlignment = HorizontalAlignment.Center
                    });

                    if (hole.NeedScan)
                    {
                        // 引物名称（期望扫码内容）
                        panel.Children.Add(new TextBlock
                        {
                            Text = hole.ExpectedPrimerLabel,
                            FontSize = 9,
                            Foreground = Brushes.DarkBlue,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            TextWrapping = TextWrapping.Wrap,
                            TextAlignment = TextAlignment.Center,
                            MaxWidth = 72
                        });

                        // 扫码状态（等待扫码 / ✔ / ✘）
                        var statusText = new TextBlock
                        {
                            Text = (string)Application.Current.FindResource("BarcodeVerifyWaiting"),
                            FontSize = 9,
                            Foreground = Brushes.Gray,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            TextWrapping = TextWrapping.Wrap,
                            TextAlignment = TextAlignment.Center,
                            MaxWidth = 72
                        };
                        panel.Children.Add(statusText);
                        _holeStatusTexts[holeIdx] = statusText;
                    }

                    border.Child = panel;
                    Grid.SetRow(border, r + 1);     // +1 跳过表头行
                    Grid.SetColumn(border, c + 1);  // +1 跳过行字母列
                    grid.Children.Add(border);

                    _holeBorders[holeIdx] = border;
                }
            }
        }

        // ══════════════════════════════════════════════
        // 队列逻辑
        // ══════════════════════════════════════════════

        private void InitQueue()
        {
            _pendingQueue.Clear();
            // 只把需要扫码的孔按 holeIndex 升序（列优先）入队
            foreach (var hole in _allHoles.Where(h => h.NeedScan))
                _pendingQueue.Enqueue(hole.HoleIndex);

            if (_pendingQueue.Count == 0)
            {
                IsConfirmed = true;
                this.Loaded += (s, e) => this.Close();
                return;
            }
            AdvanceQueue();
        }

        private void AdvanceQueue()
        {
            if (_pendingQueue.Count > 0)
            {
                _currentListIndex = _pendingQueue.Dequeue();
                HighlightCurrent();
                RefreshStatusText();
            }
            else
            {
                _currentListIndex = -1;
                IsConfirmed = true;
                this.Close(); // 全部通过，自动关闭
            }
        }

        /// <summary>高亮当前待扫孔，其余待扫孔恢复默认</summary>
        private void HighlightCurrent()
        {
            foreach (var hole in _allHoles.Where(h => h.NeedScan))
            {
                if (!_holeBorders.TryGetValue(hole.HoleIndex, out var border)) continue;

                bool isCurrent = hole.HoleIndex == _currentListIndex;
                bool isPassed = hole.ScanResult == true;
                bool isFailed = hole.ScanResult == false;

                if (isPassed)
                {
                    border.BorderBrush = Brushes.Green;
                    border.BorderThickness = new Thickness(2);
                    border.Background = new SolidColorBrush(Color.FromRgb(198, 239, 206));
                }
                else if (isFailed)
                {
                    border.BorderBrush = Brushes.Red;
                    border.BorderThickness = new Thickness(2);
                    border.Background = new SolidColorBrush(Color.FromRgb(255, 199, 199));
                }
                else if (isCurrent)
                {
                    border.BorderBrush = Brushes.OrangeRed;
                    border.BorderThickness = new Thickness(2);
                    border.Background = new SolidColorBrush(Color.FromRgb(255, 243, 220));
                }
                else
                {
                    border.BorderBrush = Brushes.SteelBlue;
                    border.BorderThickness = new Thickness(1);
                    border.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                }
            }
        }

        // ══════════════════════════════════════════════
        // 扫码处理
        // ══════════════════════════════════════════════

        private void ProcessScan(string barcode)
        {
            if (_currentListIndex < 0 || _currentListIndex >= _allHoles.Count) return;
            var hole = _allHoles[_currentListIndex];

            bool matched = string.Equals(barcode, hole.ExpectedPrimerLabel, StringComparison.Ordinal);
            hole.ScanResult = matched;

            if (_holeStatusTexts.TryGetValue(_currentListIndex, out var statusText))
            {
                if (matched)
                {
                    statusText.Text = "✔ " + barcode;
                    statusText.Foreground = Brushes.Green;
                }
                else
                {
                    statusText.Text = "✘ " + barcode + "\n"
                        + (string)Application.Current.FindResource("BarcodeVerifyExpected")
                        + hole.ExpectedPrimerLabel;
                    statusText.Foreground = Brushes.Red;
                }
            }

            HighlightCurrent();

            if (matched)
            {
                RefreshProgress();
                AdvanceQueue();
            }
            else
            {
                // 不匹配，原位重扫，状态文字已更新
                RefreshStatusText();
                RefreshProgress();
            }
        }

        // ══════════════════════════════════════════════
        // UI 刷新
        // ══════════════════════════════════════════════

        private void RefreshStatusText()
        {
            if (_currentListIndex >= 0 && _currentListIndex < _allHoles.Count)
            {
                var hole = _allHoles[_currentListIndex];
                this.TextBlockStatus.Text = string.Format(
                    (string)Application.Current.FindResource("EpRackScanPrompt"),
                    hole.HoleName, hole.ExpectedPrimerLabel);
                this.TextBlockStatus.Foreground = Brushes.DimGray;
            }
        }

        private void RefreshProgress()
        {
            int done = _allHoles.Count(h => h.NeedScan && h.ScanResult == true);
            int total = _allHoles.Count(h => h.NeedScan);
            this.TextBlockProgress.Text = string.Format(
                (string)Application.Current.FindResource("EpRackScanProgress"),
                done, total);
        }

        // ══════════════════════════════════════════════
        // Win32 钩子
        // ══════════════════════════════════════════════

        private void OnSourceInitialized(object sender, EventArgs e)
        {
            if (!_hookRegistered)
            {
                ComponentDispatcher.ThreadPreprocessMessage += OnThreadPreprocessMessage;
                _hookRegistered = true;
            }
        }

        private void OnClosed(object sender, EventArgs e)
        {
            if (_hookRegistered)
            {
                ComponentDispatcher.ThreadPreprocessMessage -= OnThreadPreprocessMessage;
                _hookRegistered = false;
            }
        }

        private void OnThreadPreprocessMessage(ref MSG msg, ref bool handled)
        {
            if (!this.IsVisible) return;
            if (msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN) return;

            int vk = msg.wParam.ToInt32();

            if (vk == VK_RETURN)
            {
                var barcode = _inputBuffer.ToString().Trim();
                _inputBuffer.Clear();
                if (!string.IsNullOrEmpty(barcode) && _currentListIndex >= 0)
                    ProcessScan(barcode);
                handled = true;
                return;
            }

            if (vk == VK_BACK)
            {
                if (_inputBuffer.Length > 0)
                    _inputBuffer.Remove(_inputBuffer.Length - 1, 1);
                handled = true;
                return;
            }

            if (_ignoredVk.Contains(vk)) return;

            var ch = VkToChar((uint)vk);
            if (ch != null)
            {
                _inputBuffer.Append(ch);
                handled = true;
            }
        }

        private static string VkToChar(uint vk)
        {
            var scanCode = MapVirtualKey(vk, 0);
            var keyboardState = new byte[256];
            if (!GetKeyboardState(keyboardState)) return null;
            var sb = new StringBuilder(4);
            int count = ToUnicode(vk, scanCode, keyboardState, sb, sb.Capacity, 0);
            return count == 1 ? sb.ToString() : null;
        }
    }
}
