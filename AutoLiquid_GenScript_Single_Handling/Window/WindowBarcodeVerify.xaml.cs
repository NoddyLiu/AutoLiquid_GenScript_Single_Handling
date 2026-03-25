using AutoLiquid_GenScript_Single_Handling.EntityCommon;
using AutoLiquid_GenScript_Single_Handling.EntityJson;
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
    ///  扫码确认盘位信息窗口（主级框）
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
        private static readonly HashSet<int> _ignoredVk = new HashSet<int>
        {
            0x10, 0x11, 0x12,
            0x14,
            0x5B, 0x5C,
            0xA0, 0xA1,
            0xA2, 0xA3,
            0xA4, 0xA5,
            0x09,
            0x1B,
        };

        // ── 数据模型 ──

        /// <summary>盘位扫码项</summary>
        public class SlotScanItem
        {
            public int TemplateIndex { get; set; }
            public string Title { get; set; }
            public string ExpectedBarcode { get; set; }
            /// <summary>是否为枪头盒（不扫码）</summary>
            public bool IsTipBox { get; set; }
            /// <summary>是否为EP管架（需弹出次级框逐管扫码）</summary>
            public bool IsEpRack { get; set; }
            /// <summary>EP管架耗材配置</summary>
            public Consumable EpRackConsumable { get; set; }
            /// <summary>
            /// EP管架各孔位对应的引物名称，按孔Index（0-based）顺序存储。
            /// 空字符串 = 该孔无对应Excel行，无需扫码。
            /// 源盘/靶盘均用此字段。
            /// </summary>
            public List<string> EpTubePrimerLabels { get; set; } = new List<string>();
            /// <summary>扫码结果：null=未扫，true=正确，false=错误</summary>
            public bool? ScanResult { get; set; }
            public string ScannedBarcode { get; set; }
        }

        // ── 字段 ──

        private readonly int _rowCount;
        private readonly int _colCount;
        private readonly int _currentRound;
        private readonly int _maxRound;

        private readonly Dictionary<int, SlotScanItem> _scanItems = new Dictionary<int, SlotScanItem>();
        private readonly Dictionary<int, Border> _slotBorders = new Dictionary<int, Border>();
        private readonly Dictionary<int, TextBlock> _slotStatusTexts = new Dictionary<int, TextBlock>();

        private readonly Queue<int> _pendingQueue = new Queue<int>();
        private int _currentScanIndex = -1;

        private readonly StringBuilder _inputBuffer = new StringBuilder();
        private bool _hookRegistered = false;

        public bool IsConfirmed { get; private set; } = false;

        // ── 构造 ──

        public WindowBarcodeVerify(List<Seq> seqList, int round, int maxRound)
        {
            InitializeComponent();

            _rowCount = ParamsHelper.Layout.RowCount;
            _colCount = ParamsHelper.Layout.ColCount;
            _currentRound = round;
            _maxRound = maxRound;

            var logoFile = AppDomain.CurrentDomain.BaseDirectory
                           + System.IO.Path.DirectorySeparatorChar + "logo.png";
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

        // ── 消息预处理 ──
        private void OnThreadPreprocessMessage(ref MSG msg, ref bool handled)
        {
            if (!this.IsVisible) return;
            if (msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN) return;

            int vk = msg.wParam.ToInt32();

            if (vk == VK_RETURN)
            {
                var barcode = _inputBuffer.ToString().Trim();
                _inputBuffer.Clear();
                if (!string.IsNullOrEmpty(barcode) && _currentScanIndex >= 0)
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

        // ── 构建扫码数据 ──

        private void BuildScanItems(List<Seq> seqList, int round)
        {
            var roundSeqs = seqList.Where(s => s.Round == round && !s.IsCmdOnly).ToList();

            // ── 枪头盒（只显示，不扫码）──
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

            // ── 源盘 ──
            var srcTick = 0;
            foreach (var seq in roundSeqs)
            {
                if (seq.IsPumpLiquid) continue;
                var srcIdx = seq.SourceTemplateIndex;

                bool isEp = seq.SourceTemplateConsumableType != null &&
                            seq.SourceTemplateConsumableType.GroupName
                                .IndexOf("ep", StringComparison.OrdinalIgnoreCase) >= 0;

                if (!_scanItems.ContainsKey(srcIdx))
                {
                    srcTick++;
                    _scanItems[srcIdx] = new SlotScanItem
                    {
                        TemplateIndex = srcIdx,
                        Title = (string)Application.Current.FindResource("TemplateSource") + srcTick,
                        ExpectedBarcode = seq.SourceTemplateName,
                        IsTipBox = false,
                        IsEpRack = isEp,
                        EpRackConsumable = isEp ? seq.SourceTemplateConsumableType : null,
                        ScanResult = null
                    };
                }

                // 收集源盘EP管架各孔的引物名称
                if (isEp && !string.IsNullOrEmpty(seq.EpTubePrimerLabel)
                    && seq.SourceHoleIndexList.Count > 0)
                {
                    int holeIdx = seq.SourceHoleIndexList[0].OriIndex;
                    var labels = _scanItems[srcIdx].EpTubePrimerLabels;
                    while (labels.Count <= holeIdx) labels.Add("");
                    labels[holeIdx] = seq.EpTubePrimerLabel;
                }
            }

            // ── 靶盘 ──
            var dstTick = 0;
            var dstSeen = new HashSet<int>();
            foreach (var seq in roundSeqs)
            {
                if (seq.IsPumpLiquid) continue;
                for (var i = 0; i < seq.TargetTemplateIndexList.Count; i++)
                {
                    var dstIdx = seq.TargetTemplateIndexList[i];

                    bool isEp = seq.TargetTemplateConsumableType != null &&
                                seq.TargetTemplateConsumableType.GroupName
                                    .IndexOf("ep", StringComparison.OrdinalIgnoreCase) >= 0;

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
                            IsEpRack = isEp,
                            EpRackConsumable = isEp ? seq.TargetTemplateConsumableType : null,
                            ScanResult = null
                        };
                    }

                    // 收集靶盘EP管架各孔的引物名称
                    if (isEp && !string.IsNullOrEmpty(seq.EpTubePrimerLabel)
                        && seq.TargetHoleIndexList.Count > i)
                    {
                        int holeIdx = seq.TargetHoleIndexList[i].OriIndex;
                        var labels = _scanItems[dstIdx].EpTubePrimerLabels;
                        while (labels.Count <= holeIdx) labels.Add("");
                        labels[holeIdx] = seq.EpTubePrimerLabel;
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
                        // 盘位标题
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
                            // 期望条形码（盘名）
                            panel.Children.Add(new TextBlock
                            {
                                Text = item.ExpectedBarcode,
                                FontSize = 11,
                                Foreground = Brushes.DarkBlue,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                TextWrapping = TextWrapping.Wrap,
                                TextAlignment = TextAlignment.Center
                            });

                            // EP管架标记
                            if (item.IsEpRack)
                            {
                                panel.Children.Add(new TextBlock
                                {
                                    Text = (string)Application.Current.FindResource("EpRackLabel"),
                                    FontSize = 10,
                                    Foreground = Brushes.DarkOrange,
                                    HorizontalAlignment = HorizontalAlignment.Center
                                });
                            }

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

                // EP管架：弹出次级扫码框
                if (_scanItems.TryGetValue(_currentScanIndex, out var item) && item.IsEpRack)
                {
                    HighlightCurrentSlot();
                    RefreshStatusText();

                    var epWnd = new WindowEpRackScan(
                        item.EpRackConsumable,
                        item.Title,
                        item.EpTubePrimerLabels)
                    {
                        Owner = this
                    };
                    epWnd.ShowDialog();

                    if (!epWnd.IsConfirmed)
                    {
                        // 次级框取消 → 主框整体取消
                        IsConfirmed = false;
                        this.Close();
                        return;
                    }

                    // 次级框全部通过，标绿盘位
                    item.ScanResult = true;
                    if (_slotBorders.TryGetValue(_currentScanIndex, out var border))
                    {
                        border.Background = new SolidColorBrush(Color.FromRgb(198, 239, 206));
                        border.BorderBrush = Brushes.Green;
                        border.BorderThickness = new Thickness(2);
                    }
                    if (_slotStatusTexts.TryGetValue(_currentScanIndex, out var statusText))
                    {
                        statusText.Text = (string)Application.Current.FindResource("EpRackScanAllPassed");
                        statusText.Foreground = Brushes.Green;
                    }

                    // 继续下一个
                    AdvanceQueue();
                    RefreshStatusText();
                }
                else
                {
                    // 普通盘位：高亮等待键盘扫码枪输入
                    HighlightCurrentSlot();
                }
            }
            else
            {
                _currentScanIndex = -1;
                CheckAllPassed();
            }
        }

        private void HighlightCurrentSlot()
        {
            foreach (var kv in _slotBorders)
            {
                kv.Value.BorderBrush = kv.Key == _currentScanIndex ? Brushes.Orange : Brushes.Gray;
                kv.Value.BorderThickness = kv.Key == _currentScanIndex ? new Thickness(2) : new Thickness(1);
            }
        }

        private void ProcessScan(string barcode)
        {
            if (_currentScanIndex < 0) return;
            if (!_scanItems.TryGetValue(_currentScanIndex, out var item)) return;

            item.ScannedBarcode = barcode;
            item.ScanResult = string.Equals(barcode, item.ExpectedBarcode, StringComparison.Ordinal);

            if (_slotBorders.TryGetValue(_currentScanIndex, out var border))
            {
                border.Background = item.ScanResult == true
                    ? new SolidColorBrush(Color.FromRgb(198, 239, 206))
                    : new SolidColorBrush(Color.FromRgb(255, 199, 199));
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
                AdvanceQueue();
                RefreshStatusText();
            }
            else
            {
                HighlightCurrentSlot();
                RefreshStatusText();
            }
        }

        // ── 检查全部通过 ──

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
                // EP管架盘位在次级框弹出期间，主框提示"EP管架逐管扫码中"
                if (item.IsEpRack)
                {
                    this.TextBlockStatus.Text = string.Format(
                        (string)Application.Current.FindResource("EpRackScanInProgress"),
                        item.Title);
                    this.TextBlockStatus.Foreground = Brushes.DarkOrange;
                }
                else
                {
                    this.TextBlockStatus.Text = string.Format(
                        (string)Application.Current.FindResource("BarcodeVerifyScanPrompt"),
                        item.Title, item.ExpectedBarcode);
                    this.TextBlockStatus.Foreground = Brushes.DimGray;
                }
            }
        }
    }
}
