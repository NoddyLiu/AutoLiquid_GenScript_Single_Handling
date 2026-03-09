using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using AutoLiquid_GenScript_Single_Handling.UserControls;
using AutoLiquid_GenScript_Single_Handling.Utils;
using AutoLiquid_Library.Enum;
using AutoLiquid_Library.Exceptions;
using AutoLiquid_Library.Utils;
using AutoLiquid_GenScript_Single_Handling.EntityJson;
using MahApps.Metro.Controls;
using ConstantsUtils = AutoLiquid_GenScript_Single_Handling.Utils.ConstantsUtils;
using Position = AutoLiquid_GenScript_Single_Handling.EntityJson.Position;

namespace AutoLiquid_GenScript_Single_Handling.Window
{
    /// <summary>
    /// 移液头设置界面
    /// </summary>
    public partial class WindowCommonSettingHead : MetroWindow
    {
        // 移液头Index
        private int mHeadIndex;

        public WindowCommonSettingHead(int headIndex)
        {
            InitializeComponent();

            this.mHeadIndex = headIndex;

            this.Loaded += OnLoaded;
            this.Closed += OnClosed;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            // 初始化控件
            InitWidget();
            // 控件事件
            ControlEvent();
        }

        private void OnClosed(object sender, EventArgs e)
        {
            MainWindow.mMainWindow.InitWindowTitle();
        }

        private void InitWidget()
        {
            ViewUtils.ShowLogo(this);

            if (this.mHeadIndex == 1)
                this.Title = (string)this.FindResource("Head2Setting");

            /**
             * 耗材类型
             */
            this.GroupBoxConsumable.Content = new ControlSettingConsumable(this.mHeadIndex);

            // /**
            //  * 移液头属性
            //  */
            // this.GroupBoxHeadAttribute.Content = new ControlSettingHeadAttribute(this.mHeadIndex);

            // /**
            //  * 取枪头
            //  */
            // this.GroupBoxTakeTip.Content = new ControlSettingTakeTip(this.mHeadIndex);

            /**
             * 退枪头
             */
            this.GroupBoxReleaseTip.Content = new ControlSettingReleaseTip(this.mHeadIndex);

            /**
             * 体积设置
             */
            this.GroupBoxVolume.Content = new ControlSettingVolume(this.mHeadIndex);

            /**
             * 速度设置
             */
            if (ParamsHelper.HeadList[this.mHeadIndex].SpeedVisible)
                this.GroupBoxSpeed.Content = new ControlSettingSpeed(this.mHeadIndex);
            else
                this.GroupBoxSpeed.Visibility = Visibility.Collapsed;

            /**
             * 验证设置
             */
            this.GroupBoxVerify.Content = new ControlSettingVerify(this.mHeadIndex);

            /**
             * 推出盘位
             */
            this.GroupBoxLaunchPlate.Content = new ControlSettingLaunchPlate(this.mHeadIndex);
            // 如果移液头Y轴不是盘移动，隐藏相关属性
            this.GroupBoxLaunchPlate.Visibility = ParamsHelper.HeadList[this.mHeadIndex].YMoveWithHead ? Visibility.Collapsed : Visibility.Visible;

            // /**
            //  * 盘位摆放
            //  */
            // this.GroupBoxPlatePutting.Content = new ControlSettingPlatePutting(this.mHeadIndex);

            // /**
            //  * 位置调试
            //  */
            // this.GroupBoxDebugPos.Content = new ControlSettingPosDebug(this.mHeadIndex);


            // 根据移液头隐藏相关复位按钮
            if (this.mHeadIndex == 0)
            {
                if (ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipAxis == EAxis.P)
                    this.BtnResetQ.Visibility = Visibility.Collapsed;
                // 如果移液头没有可变距，隐藏可变距相关属性
                if (ParamsHelper.HeadList[this.mHeadIndex].IsVariable && (ParamsHelper.HeadList[this.mHeadIndex].ChannelRow > 1 || ParamsHelper.HeadList[this.mHeadIndex].ChannelCol > 1))
                    this.BtnResetW.Visibility = Visibility.Visible;
                else
                    this.BtnResetW.Visibility = Visibility.Collapsed;

                // 喷液轴是否禁用
                if (!ParamsHelper.HeadList[this.mHeadIndex].PAvailable)
                    this.BtnResetP.Visibility = Visibility.Collapsed;
            }
            else
            {
                this.BtnResetP.Visibility = Visibility.Collapsed;
                this.BtnResetZ.Visibility = Visibility.Collapsed;

                // 喷液轴是否禁用
                if (!ParamsHelper.HeadList[this.mHeadIndex].PAvailable)
                    this.BtnResetQ.Visibility = Visibility.Collapsed;
            }
        }

        private void ControlEvent()
        {
            this.BtnResetX.Click += BtnClick;
            this.BtnResetY.Click += BtnClick;
            this.BtnResetZ.Click += BtnClick;
            this.BtnResetW.Click += BtnClick;
            this.BtnResetP.Click += BtnClick;
            this.BtnResetQ.Click += BtnClick;
            this.BtnResetAll.Click += BtnClick;

            this.BtnCommand.Click += BtnClick;
        }

        private void BtnClick(object sender, RoutedEventArgs routedEventArgs)
        {
            if (sender.Equals(this.BtnResetX))
            {
                BackgroundProcess.RunAsync(
                    delegate () { CmdHelper.Xi(this.mHeadIndex); },
                    delegate ()
                    {
                    });
            }
            else if (sender.Equals(this.BtnResetY))
            {
                BackgroundProcess.RunAsync(
                    delegate () { CmdHelper.Yi(this.mHeadIndex); },
                    delegate ()
                    {
                    });
            }
            else if (sender.Equals(this.BtnResetZ))
            {
                BackgroundProcess.RunAsync(
                    delegate () { CmdHelper.Zi(this.mHeadIndex); },
                    delegate ()
                    {
                    });
            }
            else if (sender.Equals(this.BtnResetW))
            {
                BackgroundProcess.RunAsync(
                    delegate () { CmdHelper.Wi(this.mHeadIndex); },
                    delegate ()
                    {
                    });
            }
            else if (sender.Equals(this.BtnResetP) || sender.Equals(this.BtnResetQ))
            {
                BackgroundProcess.RunAsync(
                    delegate ()
                    {
                        if (this.mHeadIndex == 0 && ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipAxis == EAxis.Q)
                            CmdHelper.Pi(1);
                        else
                            CmdHelper.Pi(this.mHeadIndex);
                    },
                    delegate ()
                    {
                    });
            }

            else if (sender.Equals(this.BtnResetAll))
            {
                BackgroundProcess.RunAsync(
                    delegate () { CmdHelper.InitMachine(false, false); },
                    delegate ()
                    {
                    });
            }
            else if (sender.Equals(this.BtnCommand))
            {
                BackgroundProcess.RunAsync(
                    delegate ()
                    {
                        var cmdStr = "";
                        Dispatcher.Invoke(() => { cmdStr = this.TextBoxCommand.Text; });
                        CmdHelper.frmDAE.DoCmd(cmdStr);
                    },
                    delegate ()
                    {
                    });
            }
        }
    }
}
