using System;
using System.IO;
using System.Linq;
using System.Text;
using MahApps.Metro.Controls;
using System.Windows;
using System.Windows.Media.Imaging;
using AutoLiquid_GenScript_Single_Handling.Utils;
using AutoUpdaterDotNET;

namespace AutoLiquid_GenScript_Single_Handling.Window
{
    /// <summary>
    /// "关于"弹出框
    /// </summary>
    public partial class WindowAbout : MetroWindow
    {
        // 更新日志文件路径
        private string logFilePath = AppDomain.CurrentDomain.BaseDirectory +
                                     Path.DirectorySeparatorChar +
                                     AutoLiquid_Library.Utils.ConstantsUtils.FILE_UPDATE_LOG;

        public WindowAbout()
        {
            InitializeComponent();

            var logoFIle = AppDomain.CurrentDomain.BaseDirectory + Path.DirectorySeparatorChar + "logo.png";
            if (File.Exists(logoFIle))
            {
                this.Icon = new BitmapImage(new Uri(logoFIle));
                this.ImageLogo.Source = new BitmapImage(new Uri(logoFIle));
            }

            try
            {
                InitSoftwareTitle();
                InitAboutInfo();
                this.LabelVersion.Content = File.ReadLines(logFilePath).First().Split(new char[] { ':', '：' }).ElementAt(1);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            // 更新日志
            UpdateLog();

            // 控件事件
            ControlEvent();
        }

        /// <summary>
        /// 软件标题
        /// </summary>
        private void InitSoftwareTitle()
        {
            var head1ChannelStr = "";
            var head1VariableStr = "";
            var head1RangeStr = "";
            var head2ChannelStr = "";
            var head2RangeStr = "";

            // 移液头1
            if (ParamsHelper.HeadList[0].IsVariable)
                head1VariableStr = (string)Application.Current.FindResource("VariableDistance");
            var channel = ParamsHelper.HeadList[0].ChannelRow * ParamsHelper.HeadList[0].ChannelCol;
            if (channel == 1)
            {
                if (!ParamsHelper.HeadList[1].Available)
                    head1ChannelStr = (string)Application.Current.FindResource("SingleChannel");
                else
                    head1ChannelStr = "1";
            }
            else
                head1ChannelStr = channel + (string)Application.Current.FindResource("Channel");
            head1RangeStr = (int)ParamsHelper.HeadList[0].HeadLiquidRange + (string)Application.Current.FindResource("Ul");

            // 移液头2
            if (ParamsHelper.HeadList[1].Available)
            {
                channel = ParamsHelper.HeadList[1].ChannelRow * ParamsHelper.HeadList[1].ChannelCol;
                head2ChannelStr = "+" + channel + (string)Application.Current.FindResource("Channel");
                head2RangeStr = "+" + (int)ParamsHelper.HeadList[1].HeadLiquidRange + (string)Application.Current.FindResource("Ul");
            }

            this.LabelSoftwareName.Content = head1ChannelStr + head1VariableStr + head2ChannelStr + (string)Application.Current.FindResource("LiquidWorkstation") + " --- " + head1RangeStr + head2RangeStr;
        }

        /// <summary>
        /// 关于信息
        /// </summary>
        private void InitAboutInfo()
        {
            var logoTxtFIle = AppDomain.CurrentDomain.BaseDirectory + Path.DirectorySeparatorChar + "logo.txt";
            // 判断是否存在logo.txt文件，如存在，就读取里面的内容
            if (File.Exists(logoTxtFIle))
            {
                var logoTxtContent = File.ReadAllText(logoTxtFIle);
                if (!logoTxtContent.Equals(""))
                    this.TextBlockAboutInfo.Text = logoTxtContent;
            }
        }

        /// <summary>
        /// 更新日志
        /// </summary>
        private void UpdateLog()
        {
            StringBuilder logContent = new StringBuilder();
            string line;
            try
            {
                StreamReader file = new StreamReader(logFilePath);
                while ((line = file.ReadLine()) != null)
                {
                    logContent.AppendLine(line);
                }

                this.TextBoxUpdateLog.Text = logContent.ToString();

                file.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private void ControlEvent()
        {
            this.BtnCheck4Update.Click += BtnClick;
            this.BtnConfirm.Click += BtnClick;
        }

        private void BtnClick(object sender, RoutedEventArgs e)
        {
            if (sender.Equals(this.BtnConfirm))
                this.Close();
            else if (sender.Equals(this.BtnCheck4Update))
            {
                AutoUpdater.ReportErrors = true;
                AutoUpdater.RunUpdateAsAdmin = false;
                AutoUpdater.Start("http://115.159.125.209/AutoLiquidSingleUpdate/UpdateInfo.xml");
            }
        }
    }
}
