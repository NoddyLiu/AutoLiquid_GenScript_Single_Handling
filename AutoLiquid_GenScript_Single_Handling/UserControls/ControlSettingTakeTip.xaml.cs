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
    /// 取枪头控件
    /// </summary>
    public partial class ControlSettingTakeTip : UserControl
    {
        // 移液头Index
        private int mHeadIndex;

        public ControlSettingTakeTip(int headIndex)
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
            this.RBtnTakeTipMethodCol.IsChecked = ParamsHelper.CommonSettingList[this.mHeadIndex].TakeTipEachCol;
            this.RBtnTakeTipMethodRow.IsChecked = !ParamsHelper.CommonSettingList[this.mHeadIndex].TakeTipEachCol;

            this.RBtnTakeTipLeft2Right.IsChecked = ParamsHelper.CommonSettingList[this.mHeadIndex].TakeTipLeft2Right;
            this.RBtnTakeTipRight2Left.IsChecked = !ParamsHelper.CommonSettingList[this.mHeadIndex].TakeTipLeft2Right;
        }

        private void ControlEvent()
        {
            this.RBtnTakeTipMethodCol.Checked += RBtnOnChecked;
            this.RBtnTakeTipMethodRow.Checked += RBtnOnChecked;

            this.RBtnTakeTipLeft2Right.Checked += RBtnOnChecked;
            this.RBtnTakeTipRight2Left.Checked += RBtnOnChecked;
        }

        private void RBtnOnChecked(object sender, RoutedEventArgs e)
        {
            if (sender.Equals(this.RBtnTakeTipMethodCol) || sender.Equals(this.RBtnTakeTipMethodRow))
                DataHelper.SaveBool(this.mHeadIndex, (bool)this.RBtnTakeTipMethodCol.IsChecked, ref ParamsHelper.CommonSettingList[this.mHeadIndex].TakeTipEachCol);
            else if (sender.Equals(this.RBtnTakeTipLeft2Right) || sender.Equals(this.RBtnTakeTipRight2Left))
                DataHelper.SaveBool(this.mHeadIndex, (bool)this.RBtnTakeTipLeft2Right.IsChecked, ref ParamsHelper.CommonSettingList[this.mHeadIndex].TakeTipLeft2Right);
        }
    }
}
