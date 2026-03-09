using AutoLiquid_GenScript_Single_Handling.EntityCommon;
using AutoLiquid_GenScript_Single_Handling.Utils;
using ControlzEx.Standard;
using MahApps.Metro.Controls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace AutoLiquid_GenScript_Single_Handling.Window
{
    /// <summary>
    ///  扫码确认盘位信息窗口
    /// </summary>
    public partial class WindowBarcodeVerify : MetroWindow
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
        // 忽略的修饰键 VK 值
        private static readonly HashSet<int> _ignoredVk = new HashSet<int>
        {
            0x10, 0x11, 0x12,       // Shift, Ctrl, Alt
            0x14,                   // CapsLock
            0x5B, 0x5C,             // LWin, RWin
            0xA0, 0xA1,             // LShift, RShift
            0xA2, 0xA3,             // LCtrl, RCtrl
            0xA4, 0xA5,             // LAlt, RAlt
            0x09,                   // Tab
            0x1B,                   // Escape
        };

        // ── 数据模型 ──

        /// <summary>盘位扫码项（枪头盒不需要扫码）</summary>
        public class SlotScanItem
        {
            /// <summary>盘位Index（0-based）</summary>
            public int TemplateIndex { get; set; }
            /// <summary>显示标题（如"源盘1"、"靶盘2"）</summary>
            public string Title { get; set; }
            /// <summary>期望条形码（即盘名）</summary>
            public string ExpectedBarcode { get; set; }
            /// <summary>是否为枪头盒（枪头盒不扫码，显示灰色）</summary>
            public bool IsTipBox { get; set; }
            /// <summary>扫码结果：null=未扫，true=正确，false=错误</summary>
            public bool? ScanResult { get; set; }
            /// <summary>已扫入的条形码文本</summary>
            public string ScannedBarcode { get; set; }
        }

        // ── 字段 ──

        private readonly int _rowCount;
        private readonly int _colCount;
        private readonly int _currentRound;
        private readonly int _maxRound;

        /// <summary>所有需要扫码的盘位项（按 TemplateIndex 索引）</summary>
        private readonly Dictionary<int, SlotScanItem> _scanItems = new Dictionary<int, SlotScanItem>();

        /// <summary>UI中每个盘位对应的Border控件</summary>
        private readonly Dictionary<int, Border> _slotBorders = new Dictionary<int, Border>();
        /// <summary>UI中每个盘位对应的状态TextBlock</summary>
        private readonly Dictionary<int, TextBlock> _slotStatusTexts = new Dictionary<int, TextBlock>();

        /// <summary>当前等待扫码的盘位Index队列（按TemplateIndex升序）</summary>
        private readonly Queue<int> _pendingQueue = new Queue<int>();

        /// <summary>当前正在等待扫码的盘位Index，-1表示全部完成</summary>
        private int _currentScanIndex = -1;

        /// <summary>USB扫码枪输入缓冲区（Enter结尾提交）</summary>
        private readonly StringBuilder _inputBuffer = new StringBuilder();

        /// <summary>ComponentDispatcher 是否已注册，防止重复注册/多次注销</summary>
        private bool _hookRegistered = false;

        /// <summary>窗口是否已确认（点击确认按钮）</summary>
        public bool IsConfirmed { get; private set; } = false;

        // ── 构造 ──

        /// <summary>
        /// 创建扫码验证窗体
        /// </summary>
        /// <param name="seqList">全部 Seq 列表</param>
        /// <param name="round">当前轮次（1-based）</param>
        /// <param name="maxRound">最大轮次</param>
        public WindowBarcodeVerify(List<Seq> seqList, int round, int maxRound)
        {
            InitializeComponent();

            _rowCount = ParamsHelper.Layout.RowCount;
            _colCount = ParamsHelper.Layout.ColCount;
            _currentRound = round;
            _maxRound = maxRound;

            var logoFile = AppDomain.CurrentDomain.BaseDirectory + System.IO.Path.DirectorySeparatorChar + "logo.png";
            if (File.Exists(logoFile))
                this.Icon = new BitmapImage(new Uri(logoFile));

            BuildScanItems(seqList, round);
            BuildLayoutGrid();
            InitQueue();
            RefreshStatusText();

            this.TextBlockRoundInfo.Text = string.Format(
                (string)Application.Current.FindResource("BarcodeVerifyRoundInfo"),
                round, maxRound);

            this.BtnConfirm.Click += (s, e) => { IsConfirmed = true; this.Close(); };
            this.BtnCancel.Click += (s, e) => { IsConfirmed = false; this.Close(); };

            // 仅在窗口初始化完成后注册一次，关闭时注销一次
            this.SourceInitialized += OnSourceInitialized;
            this.Closed += OnClosed;
        }

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

        // ── 消息预处理（在所有 HwndSourceHook 之前，包括 MahApps） ──
        private void OnThreadPreprocessMessage(ref MSG msg, ref bool handled)
        {
            // 仅当本窗口可见时处理，避免在隐藏或已关闭状态下误处理
            if (!this.IsVisible)
                return;

            // 仅关心按键按下消息
            if (msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN)
                return;

            int vk = msg.wParam.ToInt32();

            // Enter：提交缓冲区
            if (vk == VK_RETURN)
            {
                var barcode = _inputBuffer.ToString().Trim();
                _inputBuffer.Clear();
                if (!string.IsNullOrEmpty(barcode) && _currentScanIndex >= 0)
                    ProcessScan(barcode);
                handled = true;
                return;
            }

            // Backspace：删除最后一个字符
            if (vk == VK_BACK)
            {
                if (_inputBuffer.Length > 0)
                    _inputBuffer.Remove(_inputBuffer.Length - 1, 1);
                handled = true;
                return;
            }

            // 忽略修饰键本身
            if (_ignoredVk.Contains(vk))
                return;

            // 将虚拟键翻译成字符（尊重 CapsLock/Shift/键盘布局）
            var ch = VkToChar((uint)vk);
            if (ch != null)
            {
                _inputBuffer.Append(ch);
                handled = true;
            }
        }

        /// <summary>
        /// 将 Win32 Virtual-Key 转为字符，通过 GetKeyboardState + ToUnicode 实现，
        /// 正确处理 CapsLock、Shift、任意键盘布局。
        /// </summary>
        private static string VkToChar(uint vk)
        {
            var scanCode = MapVirtualKey(vk, 0 /* MAPVK_VK_TO_VSC */);
            var keyboardState = new byte[256];
            if (!GetKeyboardState(keyboardState)) return null;

            var sb = new StringBuilder(4);
            int count = ToUnicode(vk, scanCode, keyboardState, sb, sb.Capacity, 0);

            // count == 1：正常字符；count == -1：死键（忽略）；count == 0：无映射
            return count == 1 ? sb.ToString() : null;
        }

        // ── 构建扫码数据 ──

        private void BuildScanItems(List<Seq> seqList, int round)
        {
            var roundSeqs = seqList.Where(s => s.Round == round && !s.IsCmdOnly).ToList();

            // 枪头盒盘位：只显示，不扫码
            var tipTick = 0;
            foreach (var seq in roundSeqs)
            {
                if (seq.IsPumpLiquid) continue;
                var tipIdx = seq.TipTemplateIndex;
                if (!_scanItems.ContainsKey(tipIdx))
                {
                    tipTick++;
                    _scanItems[tipIdx] = new SlotScanItem
                    {
                        TemplateIndex = tipIdx,
                        Title = (string)Application.Current.FindResource("TemplateTips") + tipTick,
                        ExpectedBarcode = "",
                        IsTipBox = true,
                        ScanResult = null
                    };
                }
            }

            // 源盘
            var srcTick = 0;
            foreach (var seq in roundSeqs)
            {
                if (seq.IsPumpLiquid) continue;
                var srcIdx = seq.SourceTemplateIndex;
                if (!_scanItems.ContainsKey(srcIdx))
                {
                    srcTick++;
                    _scanItems[srcIdx] = new SlotScanItem
                    {
                        TemplateIndex = srcIdx,
                        Title = (string)Application.Current.FindResource("TemplateSource") + srcTick,
                        ExpectedBarcode = seq.SourceTemplateName,
                        IsTipBox = false,
                        ScanResult = null
                    };
                }
            }

            // 靶盘
            var dstTick = 0;
            var dstSeen = new HashSet<int>();
            foreach (var seq in roundSeqs)
            {
                if (seq.IsPumpLiquid) continue;
                for (var i = 0; i < seq.TargetTemplateIndexList.Count; i++)
                {
                    var dstIdx = seq.TargetTemplateIndexList[i];
                    if (!_scanItems.ContainsKey(dstIdx) && !dstSeen.Contains(dstIdx))
                    {
                        dstTick++;
                        dstSeen.Add(dstIdx);
                        _scanItems[dstIdx] = new SlotScanItem
                        {
                            TemplateIndex = dstIdx,
                            Title = (string)Application.Current.FindResource("TemplateTarget") + dstTick,
                            ExpectedBarcode = seq.TargetTemplateName,
                            IsTipBox = false,
                            ScanResult = null
                        };
                    }
                }
            }
        }

        // ── 构建布局网格 ──

        private void BuildLayoutGrid()
        {
            var grid = this.GridPlateLayout;
            grid.RowDefinitions.Clear();
            grid.ColumnDefinitions.Clear();
            _slotBorders.Clear();
            _slotStatusTexts.Clear();

            for (var r = 0; r < _rowCount; r++)
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            for (var c = 0; c < _colCount; c++)
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            for (var r = 0; r < _rowCount; r++)
            {
                for (var c = 0; c < _colCount; c++)
                {
                    var idx = c + r * _colCount;
                    _scanItems.TryGetValue(idx, out var item);

                    var border = new Border
                    {
                        BorderBrush = Brushes.Gray,
                        BorderThickness = new Thickness(1),
                        Margin = new Thickness(3),
                        MinHeight = 70,
                        Background = item == null
                            ? new SolidColorBrush(Color.FromRgb(245, 245, 245))
                            : item.IsTipBox
                                ? new SolidColorBrush(Color.FromRgb(220, 220, 220))
                                : new SolidColorBrush(Color.FromRgb(255, 255, 255))
                    };

                    var panel = new StackPanel
                    {
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Margin = new Thickness(4)
                    };

                    // 盘位序号
                    panel.Children.Add(new TextBlock
                    {
                        Text = (idx + 1).ToString(),
                        FontSize = 10,
                        Foreground = Brushes.DarkGray,
                        HorizontalAlignment = HorizontalAlignment.Center
                    });

                    if (item != null)
                    {
                        // 标题（枪头盒1 / 源盘1 / 靶盘1）
                        panel.Children.Add(new TextBlock
                        {
                            Text = item.Title,
                            FontSize = 11,
                            FontWeight = FontWeights.Bold,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            TextWrapping = TextWrapping.Wrap,
                            TextAlignment = TextAlignment.Center
                        });

                        if (!item.IsTipBox)
                        {
                            // 期望条形码
                            panel.Children.Add(new TextBlock
                            {
                                Text = item.ExpectedBarcode,
                                FontSize = 11,
                                Foreground = Brushes.DarkBlue,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                TextWrapping = TextWrapping.Wrap,
                                TextAlignment = TextAlignment.Center
                            });

                            // 扫码状态文本
                            var statusText = new TextBlock
                            {
                                Text = (string)Application.Current.FindResource("BarcodeVerifyWaiting"),
                                FontSize = 10,
                                Foreground = Brushes.Gray,
                                HorizontalAlignment = HorizontalAlignment.Center
                            };
                            panel.Children.Add(statusText);
                            _slotStatusTexts[idx] = statusText;
                        }
                        else
                        {
                            // 枪头盒：显示"无需扫码"
                            panel.Children.Add(new TextBlock
                            {
                                Text = (string)Application.Current.FindResource("BarcodeVerifyTipBoxNoScan"),
                                FontSize = 10,
                                Foreground = Brushes.Gray,
                                HorizontalAlignment = HorizontalAlignment.Center
                            });
                        }
                    }

                    border.Child = panel;
                    Grid.SetRow(border, r);
                    Grid.SetColumn(border, c);
                    grid.Children.Add(border);

                    if (item != null)
                        _slotBorders[idx] = border;
                }
            }
        }

        // ── 队列初始化 ──

        private void InitQueue()
        {
            _pendingQueue.Clear();
            // 只把需要扫码（非枪头盒）的盘位按TemplateIndex升序入队
            foreach (var idx in _scanItems.Keys.OrderBy(k => k))
            {
                if (!_scanItems[idx].IsTipBox)
                    _pendingQueue.Enqueue(idx);
            }
            AdvanceQueue();
        }

        private void AdvanceQueue()
        {
            if (_pendingQueue.Count > 0)
            {
                _currentScanIndex = _pendingQueue.Dequeue();
                HighlightCurrentSlot();
            }
            else
            {
                _currentScanIndex = -1;
                CheckAllPassed();
            }
        }

        private void HighlightCurrentSlot()
        {
            // 将当前待扫盘位边框高亮为橙色
            foreach (var kv in _slotBorders)
            {
                kv.Value.BorderBrush = kv.Key == _currentScanIndex
                    ? Brushes.Orange
                    : Brushes.Gray;
                kv.Value.BorderThickness = kv.Key == _currentScanIndex
                    ? new Thickness(2)
                    : new Thickness(1);
            }
        }

        private void ProcessScan(string barcode)
        {
            if (_currentScanIndex < 0) return;
            if (!_scanItems.TryGetValue(_currentScanIndex, out var item)) return;

            item.ScannedBarcode = barcode;
            item.ScanResult = string.Equals(barcode, item.ExpectedBarcode, StringComparison.Ordinal);

            // 更新UI
            if (_slotBorders.TryGetValue(_currentScanIndex, out var border))
            {
                border.Background = item.ScanResult == true
                    ? new SolidColorBrush(Color.FromRgb(198, 239, 206))   // 绿色背景
                    : new SolidColorBrush(Color.FromRgb(255, 199, 199));  // 红色背景
                border.BorderBrush = item.ScanResult == true ? Brushes.Green : Brushes.Red;
                border.BorderThickness = new Thickness(2);
            }

            if (_slotStatusTexts.TryGetValue(_currentScanIndex, out var statusText))
            {
                if (item.ScanResult == true)
                {
                    statusText.Text = "✔ " + barcode;
                    statusText.Foreground = Brushes.Green;
                }
                else
                {
                    statusText.Text = "✘ " + barcode + "\n"
                        + (string)Application.Current.FindResource("BarcodeVerifyExpected")
                        + item.ExpectedBarcode;
                    statusText.Foreground = Brushes.Red;
                }
            }

            if (item.ScanResult == true)
            {
                // 扫码正确，推进到下一个
                AdvanceQueue();
                RefreshStatusText();
            }
            else
            {
                // 扫码错误：当前盘位重新等待，继续扫
                HighlightCurrentSlot();
                RefreshStatusText();
            }
        }

        // ── 检查是否全部通过 ──

        private void CheckAllPassed()
        {
            var allPassed = _scanItems.Values
                .Where(i => !i.IsTipBox)
                .All(i => i.ScanResult == true);

            this.BtnConfirm.IsEnabled = allPassed;
            if (allPassed)
            {
                this.TextBlockStatus.Text = (string)Application.Current.FindResource("BarcodeVerifyAllPassed");
                this.TextBlockStatus.Foreground = Brushes.Green;
            }
        }

        private void RefreshStatusText()
        {
            if (_currentScanIndex >= 0 && _scanItems.TryGetValue(_currentScanIndex, out var item))
            {
                this.TextBlockStatus.Text = string.Format(
                    (string)Application.Current.FindResource("BarcodeVerifyScanPrompt"),
                    item.Title, item.ExpectedBarcode);
                this.TextBlockStatus.Foreground = Brushes.DimGray;
            }
        }
    }
}
