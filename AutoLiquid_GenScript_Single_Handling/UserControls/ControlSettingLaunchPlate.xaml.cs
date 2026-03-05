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
    /// 推出盘位控件
    /// </summary>
    public partial class ControlSettingLaunchPlate : UserControl
    {
        // 移液头Index
        private int mHeadIndex;

        public ControlSettingLaunchPlate(int headIndex)
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
            this.ControlLaunchPlateY.SetParam(this.mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex].PositionLaunchPlate.Y, EAxis.Y);
        }

        private void ControlEvent()
        {
            this.ControlLaunchPlateY.TextBoxPos.TextChanged += TextBoxOnTextChanged;
        }

        private void TextBoxOnTextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender.Equals(this.ControlLaunchPlateY.TextBoxPos))
            {
                // 同时修改移液头1和2
                DataHelper.SaveDecimal(0, this.ControlLaunchPlateY.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[0].PositionLaunchPlate.Y);
                DataHelper.SaveDecimal(1, this.ControlLaunchPlateY.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[1].PositionLaunchPlate.Y);
            }
                
        }
    }
}
