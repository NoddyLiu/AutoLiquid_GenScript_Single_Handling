using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using AutoLiquid_GenScript_Single_Handling.EntityCommon;
using AutoLiquid_GenScript_Single_Handling.EntityJson;
using AutoLiquid_GenScript_Single_Handling.Utils;
using AutoLiquid_Library.Enum;
using AutoLiquid_Library.Utils;
using MahApps.Metro.Controls;

namespace AutoLiquid_GenScript_Single_Handling.UserControls
{
    /// <summary>
    /// 枪头盒盘位模板
    /// </summary>
    public partial class ControlTemplateTip : UserControl
    {
        // 板信息
        public Template mTemplate;

        // 枪头耗材
        public Consumable mConsumableType;

        // 枪头盒使用情况二维数组（true为已使用）
        public bool[,] TipBoxUsedStatus2DArray;

        // 枪头盒是否灵活取枪头（例如8通道，可灵活取1~8根枪头）
        public bool TipBoxFlexible = false;

        // Tip盘位Index
        public int TipBoxTemplateIndex = 0;

        // 该Tip盒所属移液头Index
        public int HeadUsedIndex = 0;

        public ControlTemplateTip(Template template, Consumable consumableType)
        {
            InitializeComponent();

            mTemplate = template;
            this.mConsumableType = consumableType;

            this.Loaded += OnLoaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            InitWidget();

            ControlEvent();
        }

        private void InitWidget()
        {
            InitHoles(this.StackPanelTipsBox, Brushes.Gray);

            // Tip盒位置
            var rowCount = mTemplate.RowCount;
            var colCount = mTemplate.ColCount;
            List<string> tipsBoxPosList = new List<string>(rowCount * colCount);

            // 数字层
            var numLoop = mTemplate.A1Pos == EA1Pos.LeftTop ? colCount : rowCount;
            // 字母层
            var letterLoop = mTemplate.A1Pos == EA1Pos.LeftTop ? rowCount : colCount;
            for (var num = 0; num < numLoop; num++)
            {
                var numStr = ViewUtils.PosNumList[num];
                for (var letter = 0; letter < letterLoop; letter++)
                {
                    var letterStr = ViewUtils.PosLetterList[letter];
                    tipsBoxPosList.Add(letterStr + numStr);
                }
            }

            SplitButtonTipsBoxPos.ItemsSource = tipsBoxPosList;

            // 枪头使用情况
            TipBoxUsedStatus2DArray = new bool[rowCount, colCount];
        }

        private void ControlEvent()
        {
            // Tip位置选择
            this.SplitButtonTipsBoxPos.SelectionChanged += SplitButtonTipsBoxPosOnSelectionChanged;
        }

        private void SplitButtonTipsBoxPosOnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedIndex = 0;
            var selectedHole = (string)((SplitButton)sender).SelectedItem;
            if (selectedHole == null)
                selectedIndex = -1;
            else
                selectedIndex = ConsumableHelper.GetHoleIndex(HeadUsedIndex, this.mConsumableType, selectedHole).OriIndex;

            RefreshTemplateHolesStatus(selectedIndex);
        }

        /// <summary>
        /// Tip孔
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="borderColor"></param>
        private void InitHoles(Panel parent, Brush borderColor)
        {
            LabelTitle.Content = mTemplate.Title;

            // 判断板类型
            var colCount = mTemplate.ColCount;
            var rowCount = mTemplate.RowCount;
            var labelFontSize = 0;

            // A1位置
            var a1Pos = mTemplate.A1Pos;

            // 每个孔大小（取小值作为孔直径）
            var holeSizeHeight = parent.ActualHeight / (rowCount + 1);
            var holeSizeWidth = parent.ActualWidth / (colCount + 1);
            var holeSize = holeSizeHeight > holeSizeWidth ? holeSizeWidth : holeSizeHeight;
            labelFontSize = (int)Math.Round(holeSize, MidpointRounding.AwayFromZero) / 3;

            for (var row = 0; row < rowCount + 1; row++)
            {
                var eachRowStackPanel = ViewUtils.CreateRowStackPanel();
                for (var col = 0; col < colCount + 1; col++)
                {
                    var grid = ViewUtils.CreateGrid(holeSize);

                    // A1位置：左上（孔index顺序是同一列由上到下，再由左到右）
                    if (a1Pos == EA1Pos.LeftTop)
                    {
                        // 第一列
                        if (col == 0 && row > 0)
                        {
                            // 字母
                            var letter = ViewUtils.CreateLabel(holeSize, ViewUtils.PosLetterList[row - 1], labelFontSize);
                            grid.Children.Add(letter);
                        }
                        // 第一行
                        else if (row == 0 && col > 0)
                        {
                            // 数字
                            var num = ViewUtils.CreateLabel(holeSize, ViewUtils.PosNumList[col - 1], labelFontSize - 1);
                            grid.Children.Add(num);
                        }
                        else if (row > 0 && col > 0)
                        {
                            var holeIndex = (col - 1) * rowCount + row - 1;
                            var hole = mTemplate.Holes[holeIndex];
                            // 孔
                            hole.Circle = ViewUtils.CreateEllipse(holeSize, borderColor);
                            hole.Circle.Fill = Brushes.Yellow; // 默认为满
                            grid.Children.Add(hole.Circle);
                        }
                    }
                    // A1位置：左下（孔index顺序是同一行由左到右，再由下到上）
                    else
                    {
                        // 第一列
                        if (col == 0 && row < rowCount)
                        {
                            // 数字
                            var num = ViewUtils.CreateLabel(holeSize, ViewUtils.PosNumList[rowCount - row - 1], labelFontSize);
                            grid.Children.Add(num);
                        }
                        // 最尾行
                        else if (row == rowCount && col > 0)
                        {
                            // 字母
                            var letter = ViewUtils.CreateLabel(holeSize, ViewUtils.PosLetterList[col - 1], labelFontSize - 1);
                            grid.Children.Add(letter);
                        }
                        else if (row < rowCount && col > 0)
                        {
                            var holeIndex = (rowCount - row - 1) * colCount + (col - 1);
                            var hole = mTemplate.Holes[holeIndex];
                            // 孔
                            hole.Circle = ViewUtils.CreateEllipse(holeSize, borderColor);
                            hole.Circle.Fill = Brushes.Yellow; // 默认为满
                            grid.Children.Add(hole.Circle);
                        }
                    }

                    eachRowStackPanel.Children.Add(grid);
                }
                parent.Children.Add(eachRowStackPanel);
            }

            parent.HorizontalAlignment = HorizontalAlignment.Center;
            parent.VerticalAlignment = VerticalAlignment.Center;
        }

        /// <summary>
        /// 刷新Tips盘孔使用状态
        /// </summary>
        /// <param name="startIndex">从哪个孔开始变色（index是按列从上到下，从左往右计算）</param>
        private void RefreshTemplateHolesStatus(int startIndex)
        {
            // 行数
            var rowCount = mTemplate.RowCount;
            // 列数
            var colCount = mTemplate.ColCount;

            if (startIndex == -1)
            {
                ObjectUtils.Fill2DArray(TipBoxUsedStatus2DArray, true);
            }
            else
            {
                // 总枪头数
                var totalTipCount = rowCount * colCount;
                // 枪头是否逐列取
                var takeTipEachCol = ParamsHelper.CommonSettingList[HeadUsedIndex].TakeTipEachCol;

                // 枪头开始行index、列index
                var startRowIndex = mTemplate.A1Pos == EA1Pos.LeftTop ? startIndex % rowCount : startIndex / colCount;
                var startColIndex = mTemplate.A1Pos == EA1Pos.LeftTop ? startIndex / rowCount : startIndex % colCount;

                for (var i = 0; i < totalTipCount; i++)
                {
                    // A1左上
                    if (mTemplate.A1Pos == EA1Pos.LeftTop)
                    {
                        var currentRowIndex = i % rowCount;
                        var currentColIndex = i / rowCount;

                        // 逐列取
                        if (takeTipEachCol)
                        {
                            TipBoxUsedStatus2DArray[currentRowIndex, currentColIndex] = i < startIndex;
                        }
                        // 逐行取
                        else
                        {
                            // 置true两种情况：1、currentRowIndex > startRowIndex；2、currentRowIndex == startRowIndex && currentColIndex >= startColIndex
                            if ((currentRowIndex == startRowIndex && currentColIndex >= startColIndex || currentRowIndex > startRowIndex) && startIndex != -1)
                                TipBoxUsedStatus2DArray[currentRowIndex, currentColIndex] = false;
                            else
                                TipBoxUsedStatus2DArray[currentRowIndex, currentColIndex] = true;
                        }
                    }
                    // A1左下
                    else
                    {
                        var currentRowIndex = i / colCount;
                        var currentColIndex = i % colCount;

                        // 逐列取
                        if (takeTipEachCol)
                        {
                            // 置true两种情况：1、currentColIndex == startColIndex && currentRowIndex >= startRowIndex；2、currentColIndex > startColIndex
                            if ((currentColIndex == startColIndex && currentRowIndex >= startRowIndex || currentColIndex > startColIndex) && startIndex != -1)
                                TipBoxUsedStatus2DArray[currentRowIndex, currentColIndex] = false;
                            else
                                TipBoxUsedStatus2DArray[currentRowIndex, currentColIndex] = true;
                        }
                        // 逐行取
                        else
                        {
                            TipBoxUsedStatus2DArray[currentRowIndex, currentColIndex] = i < startIndex;
                        }
                    }
                }
            }

            // 刷新孔颜色
            RefreshTemplateHolesColor();
        }

        /// <summary>
        /// 刷新孔位颜色
        /// </summary>
        public void RefreshTemplateHolesColor()
        {
            // A1位置
            var a1Pos = ParamsHelper.CommonSettingList[this.HeadUsedIndex].A1Pos;
            // 耗材行数、列数
            var rowCount = mTemplate.RowCount;
            var colCount = mTemplate.ColCount;
            // 枪头盒使用情况
            var rowLength = TipBoxUsedStatus2DArray.GetLength(0);
            var colLength = TipBoxUsedStatus2DArray.GetLength(1);

            for (var row = 0; row < rowLength; row++)
            {
                for (var col = 0; col < colLength; col++)
                {
                    var status = TipBoxUsedStatus2DArray[row, col];

                    if (a1Pos == EA1Pos.LeftTop)
                        mTemplate.Holes[row + col * rowCount].Circle.Fill = status ? Brushes.White : Brushes.Yellow;
                    else
                        mTemplate.Holes[row * colCount + col].Circle.Fill = status ? Brushes.White : Brushes.Yellow;
                }
            }
        }
    }
}
