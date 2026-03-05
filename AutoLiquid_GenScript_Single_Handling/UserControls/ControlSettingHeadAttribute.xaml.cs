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
using AutoLiquid_GenScript_Single_Handling.Utils;
using AutoLiquid_Library.Enum;
using DataHelper = AutoLiquid_GenScript_Single_Handling.Utils.DataHelper;

namespace AutoLiquid_GenScript_Single_Handling.UserControls
{
    /// <summary>
    /// 移液头属性控件
    /// </summary>
    public partial class ControlSettingHeadAttribute : UserControl
    {
        // 移液头Index
        private int mHeadIndex;

        public ControlSettingHeadAttribute(int headIndex)
        {
            InitializeComponent();

            this.mHeadIndex = headIndex;

            this.Loaded += OnLoaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            // 初始化控件
            InitWidget();
            // 控件事件
            ControlEvent();
        }

        private void InitWidget()
        {
            this.ComboBoxChannelRow.SelectedItem = ParamsHelper.HeadList[this.mHeadIndex].ChannelRow;
            this.ComboBoxChannelCol.SelectedItem = ParamsHelper.HeadList[this.mHeadIndex].ChannelCol;
            this.ComboBoxChannelStep.SelectedItem = ParamsHelper.HeadList[this.mHeadIndex].ChannelStep;
            this.ComboBoxLiquidRange.SelectedItem = (int)ParamsHelper.HeadList[this.mHeadIndex].HeadLiquidRange;
            this.RBtnVariableNo.IsChecked = !ParamsHelper.HeadList[this.mHeadIndex].IsVariable;
            this.RBtnVariableYes.IsChecked = ParamsHelper.HeadList[this.mHeadIndex].IsVariable;

            if (ParamsHelper.HeadList[this.mHeadIndex].ChannelRow == 1 && ParamsHelper.HeadList[this.mHeadIndex].ChannelCol == 1)
            {
                this.StackPanelVariable.Visibility = Visibility.Collapsed;
            }

            // 行走逻辑
            switch (ParamsHelper.HeadList[this.mHeadIndex].WalkingLogic)
            {
                case EWalkingLogic.SameTime:
                    this.RBtnWalkingLogicSameTime.IsChecked = true;
                    break;
                case EWalkingLogic.XFirst:
                    this.RBtnWalkingLogicXFirst.IsChecked = true;
                    break;
                case EWalkingLogic.YFirst:
                    this.RBtnWalkingLogicYFirst.IsChecked = true;
                    break;
            }
        }

        private void ControlEvent()
        {
            this.ComboBoxChannelRow.SelectionChanged += ComboBoxOnSelectionChanged;
            this.ComboBoxChannelCol.SelectionChanged += ComboBoxOnSelectionChanged;
            this.ComboBoxChannelStep.SelectionChanged += ComboBoxOnSelectionChanged;
            this.ComboBoxLiquidRange.SelectionChanged += ComboBoxOnSelectionChanged;

            this.RBtnVariableNo.Checked += RBtnVariableOnChecked;
            this.RBtnVariableYes.Checked += RBtnVariableOnChecked;

            this.RBtnWalkingLogicSameTime.Checked += RBtnWalkingLogicOnChecked;
            this.RBtnWalkingLogicXFirst.Checked += RBtnWalkingLogicOnChecked;
            this.RBtnWalkingLogicYFirst.Checked += RBtnWalkingLogicOnChecked;
        }

        private void ComboBoxOnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
            if (sender.Equals(this.ComboBoxChannelRow))
            {
                var row = (int)this.ComboBoxChannelRow.SelectedItem;
                var col = (int)this.ComboBoxChannelCol.SelectedItem;
                this.StackPanelVariable.Visibility = row > 1 || col > 1 ? Visibility.Visible : Visibility.Collapsed;
                DataHelper.SaveInt(this.mHeadIndex, ((int)this.ComboBoxChannelRow.SelectedItem).ToString(), ref ParamsHelper.HeadList[this.mHeadIndex].ChannelRow, () => FileUtils.SaveHead(this.mHeadIndex, ParamsHelper.HeadList[this.mHeadIndex]));
            }
            else if (sender.Equals(this.ComboBoxChannelCol))
            {
                var row = (int)this.ComboBoxChannelRow.SelectedItem;
                var col = (int)this.ComboBoxChannelCol.SelectedItem;
                this.StackPanelVariable.Visibility = row > 1 || col > 1 ? Visibility.Visible : Visibility.Collapsed;
                DataHelper.SaveInt(this.mHeadIndex, ((int)this.ComboBoxChannelCol.SelectedItem).ToString(), ref ParamsHelper.HeadList[this.mHeadIndex].ChannelCol, () => FileUtils.SaveHead(this.mHeadIndex, ParamsHelper.HeadList[this.mHeadIndex]));
            }
            else if (sender.Equals(this.ComboBoxChannelStep))
            {
                DataHelper.SaveDecimal(this.mHeadIndex, ((decimal)this.ComboBoxChannelStep.SelectedItem).ToString(), ref ParamsHelper.HeadList[this.mHeadIndex].ChannelStep, () => FileUtils.SaveHead(this.mHeadIndex, ParamsHelper.HeadList[this.mHeadIndex]));
            }
            else if (sender.Equals(this.ComboBoxLiquidRange))
            {
                DataHelper.SaveLiquidRange((ELiquidRange)((int)this.ComboBoxLiquidRange.SelectedItem), ref ParamsHelper.HeadList[this.mHeadIndex].HeadLiquidRange, () => FileUtils.SaveHead(this.mHeadIndex, ParamsHelper.HeadList[this.mHeadIndex]));
            }
        }

        private void RBtnVariableOnChecked(object sender, RoutedEventArgs e)
        {
            if ((bool)this.RBtnVariableNo.IsChecked)
                ParamsHelper.HeadList[this.mHeadIndex].IsVariable = false;
            else if ((bool)this.RBtnVariableYes.IsChecked)
                ParamsHelper.HeadList[this.mHeadIndex].IsVariable = true;
            FileUtils.SaveHead(this.mHeadIndex, ParamsHelper.HeadList[this.mHeadIndex]);
        }

        private void RBtnWalkingLogicOnChecked(object sender, RoutedEventArgs e)
        {
            if ((bool)this.RBtnWalkingLogicXFirst.IsChecked)
                ParamsHelper.HeadList[this.mHeadIndex].WalkingLogic = EWalkingLogic.XFirst;
            else if ((bool)this.RBtnWalkingLogicYFirst.IsChecked)
                ParamsHelper.HeadList[this.mHeadIndex].WalkingLogic = EWalkingLogic.YFirst;
            else
                ParamsHelper.HeadList[this.mHeadIndex].WalkingLogic = EWalkingLogic.SameTime;
            FileUtils.SaveHead(this.mHeadIndex, ParamsHelper.HeadList[this.mHeadIndex]);
        }
    }
}
