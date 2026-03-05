using System;
using System.IO;
using System.Linq;
using System.Text;
using MahApps.Metro.Controls;
using System.Windows;
using System.Windows.Media.Imaging;
using AutoLiquid_GenScript_Single_Handling.EntityJson;
using AutoLiquid_GenScript_Single_Handling.Utils;
using AutoLiquid_Library.Enum;
using AutoLiquid_Library.Utils;

namespace AutoLiquid_GenScript_Single_Handling.Window
{
    /// <summary>
    /// 原点校准窗口
    /// </summary>
    public partial class WindowOriginCalibration : MetroWindow
    {
        // 标准机盘位原点
        decimal originalPointTemplateX = ParamsHelper.OriginalPoint.PosTemplate.X;
        decimal originalPointTemplateY = ParamsHelper.OriginalPoint.PosTemplate.Y;
        decimal originalPointTemplateZ = ParamsHelper.OriginalPoint.PosTemplate.Z;

        public WindowOriginCalibration()
        {
            InitializeComponent();

            // 初始化控件
            InitWidget();

            // 控件事件
            ControlEvent();
        }

        private void InitWidget()
        {
            ViewUtils.ShowLogo(this);

            // 根据设备id获取相应的偏移值
            var device = ParamsHelper.Offsets.Devices.FirstOrDefault(p => p.Id.Equals(CmdHelper.frmDAE.mDeviceId));
            if (null != device)
            {
                this.TextBoxOriginalPointX.Text = (originalPointTemplateX + device.OffsetTemplate.X).ToString();
                this.TextBoxOriginalPointY.Text = (originalPointTemplateY + device.OffsetTemplate.Y).ToString();
                this.TextBoxOriginalPointZ.Text = (originalPointTemplateZ + device.OffsetTemplate.Z).ToString();
            }
        }

        private void ControlEvent()
        {
            this.BtnTemplateOriginalPoint.Click += BtnOnClick;
            this.BtnCalibration.Click += BtnOnClick;
            this.BtnCancel.Click += BtnOnClick;
        }

        private void BtnOnClick(object sender, RoutedEventArgs e)
        {
            if (sender.Equals(this.BtnTemplateOriginalPoint))
            {
//                CmdHelper.XaYa(decimal.Parse(this.TextBoxOriginalPointX.Text), decimal.Parse(this.TextBoxOriginalPointY.Text), EOffsetType.No, false);
//                CmdHelper.GotoHeight(0, decimal.Parse(this.TextBoxOriginalPointZ.Text), EOffsetType.No, false);
            }
            else if (sender.Equals(this.BtnCalibration))
            {
                try
                {
                    if (originalPointTemplateX == 0 && originalPointTemplateY == 0 && originalPointTemplateZ == 0)
                    {
                        MessageBox.Show((string)Application.Current.FindResource("Prompt_Pls_Cover_Original_Point_File_First"));
                        return;
                    }
                    // 比较标准机原点数值
                    var device = ParamsHelper.Offsets.Devices.FirstOrDefault(p => p.Id.Equals(CmdHelper.frmDAE.mDeviceId));
                    if (null != device)
                    {
                        device.OffsetTemplate.X = decimal.Parse(this.TextBoxOriginalPointX.Text) - originalPointTemplateX;
                        device.OffsetTemplate.Y = decimal.Parse(this.TextBoxOriginalPointY.Text) - originalPointTemplateY;
                        device.OffsetTemplate.Z = decimal.Parse(this.TextBoxOriginalPointZ.Text) - originalPointTemplateZ;
                    }
                    // 找不到本机的校准值
                    else
                    {
                        device = new Device { Id = CmdHelper.frmDAE.mDeviceId };
                        device.OffsetTemplate.X = decimal.Parse(this.TextBoxOriginalPointX.Text) - originalPointTemplateX;
                        device.OffsetTemplate.Y = decimal.Parse(this.TextBoxOriginalPointY.Text) - originalPointTemplateY;
                        device.OffsetTemplate.Z = decimal.Parse(this.TextBoxOriginalPointZ.Text) - originalPointTemplateZ;
                        ParamsHelper.Offsets.Devices.Add(device);
                    }
                    FileUtils.SaveOffsets(ParamsHelper.Offsets);
                    // 刷新校准值
                    CmdHelper.offsetTemplate.X = device.OffsetTemplate.X;
                    CmdHelper.offsetTemplate.Y = device.OffsetTemplate.Y;
                    CmdHelper.offsetTemplate.Z = device.OffsetTemplate.Z;
                    MessageBox.Show((string)Application.Current.FindResource("Prompt_Calibration_Success"));
                }
                catch (Exception exception)
                {
                    MessageBox.Show((string)Application.Current.FindResource("Prompt_Check_Original_Point_Value"));
                }
            }
            else if (sender.Equals(this.BtnCancel))
            {
                this.Close();
            }
        }
    }
}
