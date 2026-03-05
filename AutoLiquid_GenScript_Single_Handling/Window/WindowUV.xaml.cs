using AutoLiquid_Library.Utils;
using MahApps.Metro.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using AutoLiquid_GenScript_Single_Handling.Utils;
using static AutoLiquid_GenScript_Single_Handling.MainWindow;
using Timer = System.Timers.Timer;

namespace AutoLiquid_GenScript_Single_Handling.Window
{
    /// <summary>
    /// 紫外灯控制窗体
    /// </summary>
    public partial class WindowUV : MetroWindow
    {
        // 时间间隔
        private static int mInterval = 1000;

        // 剩余时间定时器
        private Timer mTimerTimeRemain = null;
        // 剩余时间
        private double mTimeRermain = 0;

        // 门状态查询定时器
        private Timer mTimerDoorQuery = null;

        public WindowUV()
        {
            InitializeComponent();

            this.Loaded += OnLoaded;
            this.Closing += OnClosing;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            CmdHelper.UVOpen(false);
            if (ParamsHelper.IO.LightAvailable)
                CmdHelper.LightClose(false);

            // 初始化控件
            InitWidget();
            // 控件事件
            ControlEvent();
        }

        private void OnClosing(object sender, CancelEventArgs e)
        {
            CmdHelper.UVClose(false);
            if (ParamsHelper.IO.LightAvailable)
                CmdHelper.LightOpen(false);
            mTimerTimeRemain.Stop();
            mTimerDoorQuery?.Stop();
        }

        private void InitWidget()
        {
            ViewUtils.ShowLogo(this);

            mTimerTimeRemain = new Timer { Interval = mInterval, AutoReset = true, Enabled = false };
            mTimerTimeRemain.Elapsed += (sender, args) =>
            {
                // 主线程执行
                Dispatcher.Invoke(() =>
                {
                    mTimeRermain--;
                    this.ProgressBar.Value += 1;
                    RefreshRemainTime(mTimeRermain);

                    // 剩余时间为0，就关闭紫外灯
                    if (mTimeRermain == 0)
                    {
                        this.mTimerTimeRemain.Enabled = false;
                        CmdHelper.UVClose(false);
                        if (ParamsHelper.IO.LightAvailable)
                            CmdHelper.LightOpen(false);
                    }
                });
            };

            // 如果中途打开仓门，关闭紫外灯，暂停倒计时
            if (ParamsHelper.IO.DoorAvailable)
            {
                mTimerDoorQuery = new Timer { Interval = mInterval, AutoReset = true, Enabled = true };
                mTimerDoorQuery.Elapsed += (sender, args) =>
                {
                    if (mTimerTimeRemain.Enabled)
                    {
                        CmdHelper.DoorQuery(false);
                        Thread.Sleep(100);
                        // 仓门打开中
                        if (!CmdHelper.frmDAE.IsDoorClosed)
                        {
                            // 暂停倒计时
                            this.mTimerTimeRemain.Enabled = false;
                            // 关闭紫外灯
                            CmdHelper.UVClose(false);

                            CheckDoorAndShowMessageBox();
                        }
                    }
                };
            }
        }

        /// <summary>
        /// 检查门状态并弹出框提示
        /// </summary>
        private void CheckDoorAndShowMessageBox()
        {
            var promptStr = (string)Application.Current.FindResource("Prompt_Door_Is_Opened_Please_Close_And_Try_Run_UV_Again");
            if (MessageBox.Show(promptStr, (string)Application.Current.FindResource("Prompt"), MessageBoxButton.OK, MessageBoxImage.Warning) ==
                MessageBoxResult.OK)
            {
                // 查询仓门状态
                BackgroundProcess.RunAsync(() => CmdHelper.DoorQuery(false), delegate (object returnResult)
                {
                    Thread.Sleep(100);
                    // 仓门关闭
                    if (CmdHelper.frmDAE.IsDoorClosed)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            // 打开紫外灯
                            CmdHelper.UVOpen(false);
                            // 继续倒计时
                            this.mTimerTimeRemain.Enabled = true;
                        });
                    }
                    else
                    {
                        Dispatcher.Invoke(CheckDoorAndShowMessageBox);
                    }
                });
            }
        }

        private void ControlEvent()
        {
            this.BtnConfirm.Click += BtnOnClick;
            this.BtnCancel.Click += BtnOnClick;
        }

        private void BtnOnClick(object sender, RoutedEventArgs e)
        {
            if (sender.Equals(this.BtnConfirm))
            {
                try
                {
                    SetTime(int.Parse(this.TextBoxUVTimeSet.Text));
                }
                catch (Exception exception)
                {
                    MessageBox.Show((string)this.FindResource("Prompt_Data_Input_Error"));
                }

            }
            else if (sender.Equals(this.BtnCancel))
            {
                this.Close();
            }
        }

        /// <summary>
        /// 设置时长
        /// </summary>
        /// <param name="seconds">总秒数</param>
        private void SetTime(double seconds)
        {
            if (seconds <= 0)
                return;
            mTimerTimeRemain.Enabled = false;
            mTimeRermain = seconds;
            this.ProgressBar.Maximum = seconds;
            this.ProgressBar.Value = 0;
            mTimerTimeRemain.Enabled = true;
            RefreshRemainTime(seconds);

            // 开启紫外灯
            CmdHelper.UVOpen(false);
            if (ParamsHelper.IO.LightAvailable)
                CmdHelper.LightClose(false);
        }

        /// <summary>
        /// 刷新界面剩余时间
        /// </summary>
        /// <param name="seconds">总秒数</param>
        private void RefreshRemainTime(double seconds)
        {
            TimeSpan time = TimeSpan.FromSeconds(seconds);
            string result = time.ToString(@"hh\:mm\:ss");
            this.TextBlockTimeRemain.Text = result;
        }
    }
}
