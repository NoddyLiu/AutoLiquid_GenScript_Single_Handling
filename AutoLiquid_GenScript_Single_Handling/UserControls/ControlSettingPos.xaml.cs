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
using AutoLiquid_Library.Utils;
using AutoLiquid_GenScript_Single_Handling.EntityCommon;

namespace AutoLiquid_GenScript_Single_Handling.UserControls
{
    /// <summary>
    /// 位置设置控件
    /// </summary>
    public partial class ControlSettingPos : UserControl
    {
        // 移液头Index
        private int mHeadIndex;
        // 轴类型
        private EAxis mAxis;

        public ControlSettingPos()
        {
            InitializeComponent();

            ControlEvent();
        }

        /// <summary>
        /// 关联参数
        /// </summary>
        /// <param name="headIndex"></param>
        /// <param name="value"></param>
        /// <param name="axis"></param>
        public void SetParam(int headIndex, decimal value, EAxis axis)
        {
            this.mHeadIndex = headIndex;
            this.mAxis = axis;

            // 显示数值
            this.TextBoxPos.Text = value.ToString();
        }

        private void ControlEvent()
        {
            this.BtnMinus.Click += BtnTrimmingOnClick;
            this.BtnAdd.Click += BtnTrimmingOnClick;
            this.BtnExecute.Click += BtnExecuteOnClick;
        }

        private void BtnTrimmingOnClick(object sender, RoutedEventArgs e)
        {
            var trimming = ParamsHelper.Debug.IsThin ? ParamsHelper.Debug.StepThin : ParamsHelper.Debug.StepThick;
            var previousVal = decimal.Parse(this.TextBoxPos.Text.Trim());
            var currentVal = 0m;
            if (sender.Equals(this.BtnMinus))
            {
                currentVal = previousVal - trimming;
            }
            else if (sender.Equals(this.BtnAdd))
            {
                currentVal = previousVal + trimming;
            }

            this.TextBoxPos.Text = currentVal.ToString();
            DoCmd(currentVal);
        }

        private void BtnExecuteOnClick(object sender, RoutedEventArgs e)
        {
            var currentVal = decimal.Parse(this.TextBoxPos.Text.Trim());
            DoCmd(currentVal);
        }

        private void DoCmd(decimal currentVal)
        {
            if (mAxis == EAxis.X)
                CmdHelper.Xa(this.mHeadIndex, currentVal);
            else if (mAxis == EAxis.Y)
                CmdHelper.Ya(this.mHeadIndex, currentVal);
            else if (mAxis == EAxis.Z)
                CmdHelper.GotoHeight(mHeadIndex, currentVal);
            else if (this.mAxis == EAxis.P)
                CmdHelper.Pa(mHeadIndex, currentVal);
            else if (this.mAxis == EAxis.Q)
                CmdHelper.Pa(1, currentVal);
        }
    }
}
