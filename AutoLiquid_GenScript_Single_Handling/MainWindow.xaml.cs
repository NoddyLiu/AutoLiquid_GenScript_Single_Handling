using AutoLiquid_Library.Comm;
using AutoLiquid_Library.Enum;
using AutoLiquid_Library.Exceptions;
using AutoLiquid_Library.Utils;
using AutoUpdaterDotNET;
using DAERun;
using DAERun.CallBack;
using MahApps.Metro.Controls;
using Microsoft.Win32;
using Newtonsoft.Json;
using Serilog;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Serialization;
using AutoLiquid_GenScript_Single_Handling.EntityCommon;
using AutoLiquid_GenScript_Single_Handling.EntityJson;
using AutoLiquid_GenScript_Single_Handling.UserControls;
using AutoLiquid_GenScript_Single_Handling.Utils;
using AutoLiquid_GenScript_Single_Handling.Window;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using ConstantsUtils = AutoLiquid_GenScript_Single_Handling.Utils.ConstantsUtils;
using ControlTemplate = AutoLiquid_GenScript_Single_Handling.UserControls.ControlTemplate;
using DataHelper = AutoLiquid_Library.Utils.DataHelper;
using Path = System.IO.Path;
using Position = AutoLiquid_GenScript_Single_Handling.EntityJson.Position;
using Timer = System.Timers.Timer;

namespace AutoLiquid_GenScript_Single_Handling
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        public static MainWindow mMainWindow;

        // 机器当前状态
        private ERunStatus mRunStatus = ERunStatus.Initializing;

        // 所有任务名称
        private List<string> excelFilesNameList;
        // 导入的任务名称
        private string excelFilesNameImported;

        // 程序是否运行中（包括运行中和停止两种情况）
        public bool isRunning;

        // 运行弹出框可拖拽
        private Point loadingStartPoint;
        private Vector loadingStartOffset;

        // 移液指令信息
        private List<Seq> seqList = new List<Seq>();

        // 枪头盘位（key：所在盘位index）
        public static Dictionary<int, ControlTemplateTip> tipTemplateDict = new Dictionary<int, ControlTemplateTip>();

        // 源盘盘位（key：所在盘位index）
        public static Dictionary<int, ControlTemplateCommon> sourceTemplateDict =
            new Dictionary<int, ControlTemplateCommon>();

        // 靶盘盘位（key：所在盘位index）
        public static Dictionary<int, ControlTemplateCommon> targetTemplateDict =
            new Dictionary<int, ControlTemplateCommon>();

        // 可用盘位
        private List<Grid> templateCanUse = new List<Grid>();

        /**
         * 紧急状态
         */
        // 当前指令是否被急停拦截了
        public static bool mIsEmergencyStopIntercept;
        // 紧急停止是否再次关闭
        public static bool mIsEmergencyStopClosedAgain;
        // 急停最后一条拦截的指令
        public static string mEmergencyStopLastCmd = "";

        // 串口通信
        private SerialPort mSerialPort;

        // 窗体右上角关闭按钮是否可用
        private bool mWindowCloseDisable = false;

        // 网络连接定时器（每n小时发送一个xs0指令，避免断开连接。运行期间不发送）
        private System.Timers.Timer connectTimer;

        public MainWindow()
        {
            InitializeComponent();

            this.Loaded += OnLoaded;
            this.Closing += OnClosing;
            this.Closed += OnClosed;
        }

        private void OnLoaded(object sender, RoutedEventArgs routedEventArgs)
        {
            mMainWindow = this;

            // 加载所有参数
            if (!ParamsHelper.LoadAllParams())
                MessageBox.Show((string)this.FindResource("Prompt_Params_Parse_Error_Check_Log_For_Help"));

            // 网络通信
            CanNet();

            // 初始化控件
            InitWidget();

            // 控件事件
            ControlEvent();

            // 校准偏移值
            Calibration();

            // 外部设备串口控制
            SerialPort();
        }

        private void OnClosing(object sender, CancelEventArgs e)
        {
            e.Cancel = mWindowCloseDisable;
        }

        private void OnClosed(object sender, EventArgs eventArgs)
        {
            if (ParamsHelper.IO.DoorAvailable)
                // 关闭开门停止功能
                CmdHelper.DoorStopClose(false);
            if (ParamsHelper.IO.EmergencyStopAvailable)
                // 关闭紧急停止功能
                CmdHelper.EmergencyStopClose(false);

            if (ParamsHelper.IO.SerialPortAvailable)
            {
                if (mSerialPort != null && mSerialPort.IsOpen)
                    mSerialPort.Close();
            }

            CmdHelper.frmDAE.Disconnect();
            CmdHelper.frmDAE.Close();

            connectTimer?.Close();
        }

        /// <summary>
        /// 设置网络通信
        /// </summary>
        private void CanNet()
        {
            if (!NetUtils.TestNet(ParamsHelper.IO.IP))
            {
                MessageBox.Show((string)this.FindResource("Prompt_Host_Disconnect_Pls_Check_Again"));
                SetLoadingStatus(ERunStatus.Stop);
                // 打开红指示灯
                if (ParamsHelper.IO.WarningLightAvailable)
                    CmdHelper.WarningLightOn(EWarningLight.Red, false);
                return;
            }

            CmdHelper.frmDAE.isCanNetSuccess = true;
            CmdHelper.frmDAE.IsShowFormTip = false;
            try
            {
                SetLoadingStatus(ERunStatus.Initializing);

                // 定时器
                Timer();

                BackgroundProcess.RunAsync(() => CmdHelper.InitMachineAndEasy2Put(false, true), delegate (object returnResult)
                {
                    SetLoadingStatus(ERunStatus.Stop);
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show((string)this.FindResource("Prompt_Init_Error_Pls_Check_Again"));
                LogHelper.Error(ex.StackTrace);
                SetLoadingStatus(ERunStatus.Stop);
                // 打开红指示灯
                if (ParamsHelper.IO.WarningLightAvailable)
                    CmdHelper.WarningLightOn(EWarningLight.Red, false);
            }
        }

        /// <summary>
        /// 定时器
        /// </summary>
        private void Timer()
        {
            // 网络连接定时器
            connectTimer = new System.Timers.Timer();
            connectTimer.Interval = 60 * 60 * 1000 * 0.5; // 半小时
            connectTimer.Elapsed += new System.Timers.ElapsedEventHandler(ConnectTimer);
            connectTimer.AutoReset = true;
            connectTimer.Enabled = true;
        }

        public void ConnectTimer(object source, System.Timers.ElapsedEventArgs e)
        {
            if (mRunStatus == ERunStatus.Pause || mRunStatus == ERunStatus.Stop || mRunStatus == ERunStatus.Countdown)
            {
                LogHelper.Info((string)Application.Current.FindResource("Prompt_Connect_Check"), "");
                CmdHelper.Xs(0, 0);
            }
        }

        /// <summary>
        /// 校准偏移值
        /// </summary>
        private void Calibration()
        {
            /**
             * 根据设备id获取相应的位置偏移值（校准）
             */
            var thread = new Thread(() =>
            {
                // 查询设备id
                CmdHelper.QueryDeviceId();
                Thread.Sleep(300);
                var device = ParamsHelper.Offsets.Devices.FirstOrDefault(p => p.Id.Equals(CmdHelper.frmDAE.mDeviceId));
                if (null != device)
                {
                    CmdHelper.offsetTemplate.X = device.OffsetTemplate.X;
                    CmdHelper.offsetTemplate.Y = device.OffsetTemplate.Y;
                    CmdHelper.offsetTemplate.Z = device.OffsetTemplate.Z;
                }
                // 获取不了设备id，默认用最后一组数据（即最新数据）
                else
                {
                    var deviceNewest = ParamsHelper.Offsets.Devices.LastOrDefault();
                    if (null != deviceNewest)
                    {
                        CmdHelper.offsetTemplate.X = deviceNewest.OffsetTemplate.X;
                        CmdHelper.offsetTemplate.Y = deviceNewest.OffsetTemplate.Y;
                        CmdHelper.offsetTemplate.Z = deviceNewest.OffsetTemplate.Z;
                    }
                }
            });
            thread.IsBackground = true;
            thread.Start();
        }


        #region 串口通信
        /// <summary>
        /// 外部设备串口控制
        /// </summary>
        private void SerialPort()
        {
            if (ParamsHelper.IO.SerialPortAvailable)
            {
                try
                {
                    mSerialPort = new SerialPort(ParamsHelper.IO.SeialPort, 9600) { DtrEnable = true };
                    mSerialPort.DataReceived += new SerialDataReceivedEventHandler(SerialPortOnDataReceived);
                    mSerialPort.Open();
                }
                catch (Exception e)
                {
                    LogHelper.Error((string)this.FindResource("Prompt_Serial_Port_Error") + "：" + e.StackTrace);
                }
            }
        }

        private void SerialPortOnDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            var serialDevice = sender as SerialPort;
            // 收取完整帧
            while (serialDevice.BytesToRead != 8)
            {
                Thread.Sleep(10);
            }
            var byteCount = serialDevice.BytesToRead;
            var bytes = new byte[byteCount];
            serialDevice.Read(bytes, 0, byteCount);
            // 打印串口接收数据
            string hexString = String.Join(" ", bytes.Select(b => b.ToString("X2")));
            LogHelper.Info((string)this.FindResource("Prompt_Serial_Data_Receiver"), hexString);
            // 构建帧
            BaseFrame baseFrame = new BaseFrame();
            try
            {
                baseFrame.CopyFromBuffer(bytes);
                // 判断校验码
                if (DataHelper.CheckFrameCodeValidate(baseFrame))
                {
                    // 控制设备（启停程序）
                    if (baseFrame.FrameData.CMDCode == Code.CMD_CONTROL && baseFrame.FrameData.ConcreteData[0] == Code.DEVICE_PROGRAM)
                    {
                        // 设备或程序编号（默认为1）
                        var deviceProgramNo = baseFrame.FrameData.ConcreteData[2];

                        // 通信失败
                        // if (!CmdHelper.frmDAE.isCanNetSuccess)
                        // {
                        //     var validData = new ValidData(Code.CMD_CONTROL, new byte[] { Code.DEVICE_PROGRAM, Code.ERROR_COMM, deviceProgramNo });
                        //     SerialPortSender(serialDevice, validData);
                        // }
                        // // 没有可执行的程序
                        // else if (excelFilesNameList.Count < deviceProgramNo || deviceProgramNo == 0)
                        // {
                        //     var validData = new ValidData(Code.CMD_CONTROL, new byte[] { Code.DEVICE_PROGRAM, Code.ERROR_NO_PROGRAM, deviceProgramNo });
                        //     SerialPortSender(serialDevice, validData);
                        // }
                        // else
                        // {
                        // 启动程序
                        if (baseFrame.FrameData.ConcreteData[1] == 0x01)
                        {
                            if (!isRunning)
                            {
                                // 根据设备或程序编号判断需要导入运行哪个excel文件
                                string rootPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                                excelFilesNameImported = excelFilesNameList.ElementAt(deviceProgramNo - 1);
                                var filePath = rootPath + @"\" + excelFilesNameImported;
                                ParseDataFromExcel(filePath);

                                bool checkBoxReplaceTipOverRange = false;
                                Dispatcher.Invoke(() =>
                                {
                                    checkBoxReplaceTipOverRange = (bool)this.CheckBoxReplaceTipOverRange.IsChecked;
                                });
                                RunProcedure(checkBoxReplaceTipOverRange, false, new SerialPortCompletedCallBack(serialDevice, deviceProgramNo));
                            }
                        }
                        // 停止程序
                        else if (baseFrame.FrameData.ConcreteData[1] == 0x00)
                        {
                            if (isRunning)
                            {
                                // 模拟点击停止
                                mMainWindow.Dispatcher.Invoke(() =>
                                {
                                    // 模拟手动点击“暂停”
                                    SetLoadingStatus(ERunStatus.Pause, false);
                                    Thread.Sleep(100);
                                    // 模拟手动点击“停止”
                                    SetLoadingStatus(ERunStatus.Stop, false);
                                });
                            }
                            var validData = new ValidData(Code.CMD_CONTROL, new byte[] { Code.DEVICE_PROGRAM, 0x00, deviceProgramNo });
                            SerialPortSender(serialDevice, validData);
                        }
                        // }
                    }
                    // 查询状态（移液工作站）
                    else if (baseFrame.FrameData.CMDCode == Code.CMD_STATUS && baseFrame.FrameData.ConcreteData[0] == Code.DEVICE_PROGRAM)
                    {
                        // 帧有效数据
                        ValidData validData;

                        // 通信失败
                        if (!CmdHelper.frmDAE.isCanNetSuccess)
                            validData = new ValidData(Code.CMD_STATUS, new byte[] { Code.DEVICE_PROGRAM, Code.ERROR_COMM, 0 });
                        else
                        {
                            // 运行中
                            if (isRunning)
                                validData = new ValidData(Code.CMD_STATUS, new byte[] { Code.DEVICE_PROGRAM, 0x01, 0 });
                            // 待机中
                            else
                                validData = new ValidData(Code.CMD_STATUS, new byte[] { Code.DEVICE_PROGRAM, 0x00, 0 });
                        }
                        SerialPortSender(serialDevice, validData);
                    }
                }
                else
                    LogHelper.Error((string)this.FindResource("Prompt_Check_Code_Error"));
            }
            catch (Exception ex)
            {
                LogHelper.Error((string)this.FindResource("Prompt_Frame_Error") + "：" + ex.StackTrace);
            }

        }

        /// <summary>
        /// 串口发送数据
        /// </summary>
        /// <param name="serialDevice"></param>
        /// <param name="responseFrame"></param>
        private static void SerialPortSender(SerialPort serialDevice, ValidData validData)
        {
            BaseFrame responseFrame = new BaseFrame(validData);
            serialDevice.Write(responseFrame.TotalData, 0, responseFrame.TotalData.Length);
            // 打印串口响应数据
            string hexString = String.Join(" ", responseFrame.TotalData.Select(b => b.ToString("X2")));
            LogHelper.Info((string)Application.Current.FindResource("Prompt_Serial_Data_Sender"), hexString);
        }

        /// <summary>
        /// 串口执行完成回调
        /// </summary>
        public class SerialPortCompletedCallBack : MyCallBack
        {
            private SerialPort mSerialPort;
            // 程序编号
            private byte mProgramNo;

            public SerialPortCompletedCallBack(SerialPort serialDevice, byte programNo)
            {
                mSerialPort = serialDevice;
                mProgramNo = programNo;
            }

            public override void Callback()
            {
                var validData = new ValidData(Code.CMD_CONTROL, new byte[] { Code.DEVICE_PROGRAM, 0x00, mProgramNo });
                SerialPortSender(mSerialPort, validData);
            }
        }
        #endregion

        /// <summary>
        /// 设置loading框状态
        /// </summary>
        /// <param name="status"></param>
        /// <param name="needControlIO">控制IO</param>
        public void SetLoadingStatus(ERunStatus status, bool needControlIO = true)
        {
            switch (status)
            {
                case ERunStatus.Initializing:
                    if (needControlIO)
                    {
                        // 关闭开门停止功能
                        if (ParamsHelper.IO.DoorAvailable)
                            CmdHelper.DoorStopClose(false);
                        // 关闭紧急停止功能
                        if (ParamsHelper.IO.EmergencyStopAvailable)
                            CmdHelper.EmergencyStopClose(false);
                    }

                    this.LayoutRunningLoading.Visibility = Visibility.Visible;
                    mWindowCloseDisable = true;
                    this.ProgressRingLoading.Visibility = Visibility.Visible;
                    this.LabelPrompt.Content = (string)this.FindResource("Prompt_Initializing");
                    this.DockPanelAllBtn.Visibility = Visibility.Collapsed;
                    this.StackPanelIntervalRemain.Visibility = Visibility.Collapsed;

                    // 按钮不可用
                    MainUiButtonAvailable(false);

                    CmdHelper.isManualPause = false;
                    CmdHelper.isManualStop = false;
                    isRunning = false;
                    break;
                case ERunStatus.Pause:
                    this.LayoutRunningLoading.Visibility = Visibility.Visible;
                    mWindowCloseDisable = true;
                    this.ProgressRingLoading.Visibility = Visibility.Collapsed;
                    this.LabelPrompt.Content = (string)this.FindResource("Prompt_Manual_Pause");
                    this.BtnPause.Visibility = Visibility.Collapsed;
                    this.StackPanelStopOrContinue.Visibility = Visibility.Visible;

                    // 按钮不可用
                    MainUiButtonAvailable(false);

                    // 更新运行进度条
                    this.ControlProgressBar.SetProgressBar(ERunStatus.Pause);

                    CmdHelper.isManualPause = true;
                    CmdHelper.isManualStop = false;
                    isRunning = true;

                    // 以下代码一定要放在isManualPause和isManualStop下面，避免线程冲突同时下发指令导致下位机指令被冲
                    if (needControlIO)
                    {
                        // 打开指示黄灯
                        if (ParamsHelper.IO.WarningLightAvailable)
                        {
                            CmdHelper.ResetAllWarningLight(false);
                            Thread.Sleep(30);
                            CmdHelper.WarningLightOn(EWarningLight.Yellow, false);
                        }
                    }
                    break;
                case ERunStatus.Stop:
                    if (needControlIO)
                    {
                        // 关闭开门停止功能
                        if (ParamsHelper.IO.DoorAvailable)
                            CmdHelper.DoorStopClose(false);
                        // 关闭紧急停止功能
                        if (ParamsHelper.IO.EmergencyStopAvailable)
                            CmdHelper.EmergencyStopClose(false);

                        // 关闭指示灯
                        if (ParamsHelper.IO.WarningLightAvailable)
                            CmdHelper.ResetAllWarningLight(false);

                        // 关闭风扇
                        if (ParamsHelper.IO.FanAvailable)
                            CmdHelper.FanClose(false);
                    }

                    this.LayoutRunningLoading.Visibility = Visibility.Collapsed;
                    mWindowCloseDisable = false;

                    // 按钮可用
                    MainUiButtonAvailable(true);

                    // 更新运行进度条
                    this.ControlProgressBar.SetProgressBar(ERunStatus.Stop);

                    CmdHelper.isManualPause = true;
                    CmdHelper.isManualStop = true;
                    isRunning = false;
                    break;
                case ERunStatus.Continue:
                case ERunStatus.Running:
                    // 继续
                    if (status == ERunStatus.Continue)
                        // 更新运行进度条
                        this.ControlProgressBar.SetProgressBar(ERunStatus.Continue);
                    // 运行
                    else
                    {
                        // 更新运行进度条
                        this.ControlProgressBar.SetProgressBar(0.0);
                        this.ControlProgressBar.SetProgressBar(ERunStatus.Running);

                        this.StackPanelIntervalRemain.Visibility = Visibility.Collapsed;
                    }

                    if (needControlIO)
                    {
                        // 打开指示绿灯
                        if (ParamsHelper.IO.WarningLightAvailable)
                        {
                            CmdHelper.ResetAllWarningLight(false);
                            Thread.Sleep(30);
                            CmdHelper.WarningLightOn(EWarningLight.Green, false);
                        }

                        if (status == ERunStatus.Running)
                        {
                            // 打开开门停止功能
                            if (ParamsHelper.IO.DoorAvailable)
                                CmdHelper.DoorStopOpen(false);
                            // 打开紧急停止功能
                            if (ParamsHelper.IO.EmergencyStopAvailable)
                                CmdHelper.EmergencyStopOpen(false);

                            // 打开风扇
                            if (ParamsHelper.IO.FanAvailable)
                                CmdHelper.FanOpen(false);
                        }
                    }

                    this.LayoutRunningLoading.Visibility = Visibility.Visible;
                    mWindowCloseDisable = true;
                    this.ProgressRingLoading.Visibility = Visibility.Visible;
                    this.LabelPrompt.Content = (string)this.FindResource("Prompt_Running_Pls_Wait");
                    this.DockPanelAllBtn.Visibility = Visibility.Visible;
                    this.BtnPause.Visibility = Visibility.Visible;
                    this.StackPanelStopOrContinue.Visibility = Visibility.Collapsed;

                    // 按钮不可用
                    MainUiButtonAvailable(false);

                    CmdHelper.isManualPause = false;
                    CmdHelper.isManualStop = false;
                    isRunning = true;
                    break;
            }

            mRunStatus = status;
        }

        /// <summary>
        /// 根目录读取Excel
        /// </summary>
        private void ReadExcelFromRootPath()
        {
            string rootPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            excelFilesNameList = Directory.EnumerateFiles(rootPath, "*.*", SearchOption.TopDirectoryOnly)
                .Where(s => s.IndexOfAny(new char[] { '~', '$' }) < 0 && (s.EndsWith(".xlsx") || s.EndsWith(".xls"))).Select(Path.GetFileName).OrderBy(f => f).ToList();


            // 如果只有一个excel文件，默认自动导入
            var excelFilesNameListCount = excelFilesNameList.Count();
            if (excelFilesNameListCount == 1)
            {
                excelFilesNameImported = excelFilesNameList.ElementAt(0);
                var filePath = rootPath + @"\" + excelFilesNameImported;
                BackgroundProcess.RunAsync(() => ParseDataFromExcel(filePath), delegate (object returnResult)
                {
                });
            }

            // 生成任务栏
            for (var i = 0; i < excelFilesNameListCount; i++)
            {
                var fileName = excelFilesNameList.ElementAt(i);
                var menuItem = new MenuItem { Header = Path.GetFileNameWithoutExtension(fileName) };
                menuItem.Click += (sender, args) =>
                {
                    excelFilesNameImported = fileName;
                    ImportExcel(rootPath + @"\" + fileName);
                };
                this.MenuItemTask.Items.Add(menuItem);
            }
        }

        /// <summary>
        /// 显示窗体标题
        /// </summary>
        public void InitWindowTitle()
        {
            var head1ChannelStr = "";
            var head1VariableStr = "";
            var head1RangeStr = "";
            var head2ChannelStr = "";
            var head2RangeStr = "";

            // 移液头1
            if (ParamsHelper.HeadList[0].IsVariable)
                head1VariableStr = (string)mMainWindow.FindResource("VariableDistance");
            var channel = ParamsHelper.HeadList[0].ChannelRow * ParamsHelper.HeadList[0].ChannelCol;
            if (channel == 1)
            {
                if (!ParamsHelper.HeadList[1].Available)
                    head1ChannelStr = (string)mMainWindow.FindResource("SingleChannel");
                else
                    head1ChannelStr = "1";
            }
            else
                head1ChannelStr = channel + (string)mMainWindow.FindResource("Channel");
            head1RangeStr = (int)ParamsHelper.HeadList[0].HeadLiquidRange + (string)mMainWindow.FindResource("Ul");

            // 移液头2
            if (ParamsHelper.HeadList[1].Available)
            {
                channel = ParamsHelper.HeadList[1].ChannelRow * ParamsHelper.HeadList[1].ChannelCol;
                head2ChannelStr = "+" + channel + (string)mMainWindow.FindResource("Channel");
                head2RangeStr = "+" + (int)ParamsHelper.HeadList[1].HeadLiquidRange + (string)mMainWindow.FindResource("Ul");
            }

            mMainWindow.Title = head1ChannelStr + head1VariableStr + head2ChannelStr + (string)mMainWindow.FindResource("LiquidWorkstation") + " --- " + head1RangeStr + head2RangeStr;
        }

        /// <summary>
        /// 主界面按钮是否可点击
        /// </summary>
        /// <param name="available"></param>
        private void MainUiButtonAvailable(bool available)
        {
            this.MenuBar.IsEnabled = available;
            this.GridTemplateContainer.IsEnabled = available;
            this.BottomBar.IsEnabled = available;
        }

        /// <summary>
        /// 弹出框控制按钮是否可点击
        /// </summary>
        /// <param name="available"></param>
        private void DockPanelAllBtnAvailable(bool available)
        {
            this.DockPanelAllBtn.IsEnabled = available;
        }

        private void InitWidget()
        {
            ViewUtils.ShowLogo(this);

            // 权限是否启用
            if (!ParamsHelper.Permission.Available)
                this.MenuItemPermission.Visibility = Visibility.Collapsed;

            /**
             * 添加可用盘位
             */
            // 动态生成盘位分布
            var rowCount = ParamsHelper.Layout.RowCount;
            var colCount = ParamsHelper.Layout.ColCount;
            for (var row = 0; row < rowCount * ConstantsUtils.TemplateOccupyGridSpan; row++)
            {
                this.GridTemplateContainerOnlyNumber.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                this.GridTemplateContainer.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            }
            for (var col = 0; col < colCount * ConstantsUtils.TemplateOccupyGridSpan; col++)
            {
                this.GridTemplateContainerOnlyNumber.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                this.GridTemplateContainer.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            }
            for (var row = 0; row < rowCount; row++)
            {
                for (var col = 0; col < colCount; col++)
                {
                    var templateIndex = col + row * colCount;

                    // 只显示盘符的盘位
                    var templateOnlyNum = new ControlTemplateOnlyNumber(templateIndex);
                    Grid.SetRow(templateOnlyNum, row * ConstantsUtils.TemplateOccupyGridSpan);
                    Grid.SetRowSpan(templateOnlyNum, ConstantsUtils.TemplateOccupyGridSpan);
                    Grid.SetColumn(templateOnlyNum, col * ConstantsUtils.TemplateOccupyGridSpan);
                    Grid.SetColumnSpan(templateOnlyNum, ConstantsUtils.TemplateOccupyGridSpan);
                    this.GridTemplateContainerOnlyNumber.Children.Add(templateOnlyNum);
                    // 真正填充耗材的盘位（后期根据耗材占用盘位数动态改变Span值）
                    var template = new UserControls.ControlTemplate(templateIndex);
                    Grid.SetRow(template, 0);
                    Grid.SetRowSpan(template, ConstantsUtils.TemplateOccupyGridSpan);
                    Grid.SetColumn(template, 0);
                    Grid.SetColumnSpan(template, ConstantsUtils.TemplateOccupyGridSpan);
                    this.GridTemplateContainer.Children.Add(template);

                    templateCanUse.Add(template.Template);
                }
            }

            // 根目录读取Excel表格
            Task.Delay(1000).ContinueWith(t =>
            {
                Dispatcher.Invoke(ReadExcelFromRootPath);
            });

            // 初始化标题
            InitWindowTitle();

            // 如果有开门 或者 紧急停止，就拦截指令
            if (ParamsHelper.IO.DoorAvailable || ParamsHelper.IO.EmergencyStopAvailable)
            {
                // 设置指令被拦截回调
                CmdHelper.frmDAE.SetInterceptedCallBack(new CmdInterceptedCallBack());
            }

            // 初始化IO设备按钮
            InitIOWidget();

            // 更新运行进度条
            this.ControlProgressBar.SetProgressBar(ERunStatus.Initializing);
        }

        /// <summary>
        /// 初始化IO设备按钮
        /// </summary>
        private void InitIOWidget()
        {
            // 推出盘位
            if (ParamsHelper.HeadList[0].YMoveWithHead)
                this.BtnLaunchPlate.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// 指令被拦截回调
        /// </summary>
        public class CmdInterceptedCallBack : MyCallBack
        {
            public override void Callback()
            {
                if (mMainWindow.isRunning)
                {
                    mIsEmergencyStopIntercept = true;

                    // 马上停止下一个指令执行
                    CmdHelper.isManualPause = true;
                    CmdHelper.isManualStop = false;

                    mMainWindow.Dispatcher.Invoke(() =>
                    {
                        // 模拟手动点击“暂停”
                        mMainWindow.SetLoadingStatus(ERunStatus.Pause, true);
                        mMainWindow.CheckEmergencyStopAndShowMessageBox();
                    });
                }
            }
        }

        /// <summary>
        /// 检查急停状态并弹出框提示
        /// </summary>
        private void CheckEmergencyStopAndShowMessageBox()
        {
            // 弹窗存在禁止点击“继续”等按钮
            DockPanelAllBtnAvailable(false);
            var promptStr = (string)Application.Current.FindResource("Prompt_EmergencyStop_Is_Opened_Please_Init_First");
            if (MessageBox.Show(promptStr, (string)Application.Current.FindResource("Prompt"), MessageBoxButton.OK, MessageBoxImage.Warning) ==
                MessageBoxResult.OK)
            {
                // 查询仓门和急停状态
                BackgroundProcess.RunAsync(() => CmdHelper.DoorAndEmergencyStopQuery(false), delegate (object returnResult)
                {
                    Thread.Sleep(100);
                    var resultDoor = ParamsHelper.IO.DoorAvailable ? CmdHelper.frmDAE.IsDoorClosed : true;
                    var resultEmergencyStop = ParamsHelper.IO.EmergencyStopAvailable ? CmdHelper.frmDAE.IsEmergencyStopClosed : true;
                    if (resultDoor && resultEmergencyStop)
                    {
                        mMainWindow.Dispatcher.Invoke(() =>
                        {
                            mIsEmergencyStopClosedAgain = true;
                            mIsEmergencyStopIntercept = false;
                            // 恢复“继续”等按钮可点击
                            DockPanelAllBtnAvailable(true);
                        });
                    }
                    else
                    {
                        mMainWindow.Dispatcher.Invoke(() =>
                        {
                            mIsEmergencyStopClosedAgain = false;
                            CheckEmergencyStopAndShowMessageBox();
                        });
                    }
                });
            }
        }

        private void ControlEvent()
        {
            // 键盘事件
            this.KeyDown += OnKeyDown;

            // 菜单栏
            this.MenuItemUserSettingHead.Click += MenuItemOnClick;
            this.MenuItemModifyPwd.Click += MenuItemOnClick;
            this.MenuItemAbout.Click += MenuItemOnClick;

            // 运行
            this.BtnRun.Click += BtnOnClick;

            // 复位
            this.BtnReset.Click += BtnOnClick;

            // 推出盘位
            this.BtnLaunchPlate.Click += BtnOnClick;

            // 运行弹出框
            this.BtnPause.Click += BtnOnClick;
            this.BtnStop.Click += BtnOnClick;
            this.BtnContinue.Click += BtnOnClick;

            // 导入
            this.BtnImportData.Click += BtnOnClick;

            // 运行进度框可拖拽
            this.LayoutRunningLoading.MouseDown += (sender, args) =>
            {
                loadingStartPoint = args.GetPosition(this.mainWindow);
                loadingStartOffset = new Vector(LayoutLoadingRenderTransform.X, LayoutLoadingRenderTransform.Y);
                this.LayoutRunningLoading.CaptureMouse();
            };
            this.LayoutRunningLoading.MouseMove += (sender, args) =>
            {
                if (this.LayoutRunningLoading.IsMouseCaptured)
                {
                    Vector offset = Point.Subtract(args.GetPosition(this.mainWindow), loadingStartPoint);

                    LayoutLoadingRenderTransform.X = loadingStartOffset.X + offset.X;
                    LayoutLoadingRenderTransform.Y = loadingStartOffset.Y + offset.Y;
                }
            };
            this.LayoutRunningLoading.MouseUp += (sender, args) => { this.LayoutRunningLoading.ReleaseMouseCapture(); };
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space) // 空格，暂停 or 继续
            {
                if (isRunning)
                {
                    if (CmdHelper.isManualPause && !CmdHelper.isManualStop) // 暂停中按空格键
                    {
                        // 继续
                        this.BtnContinue.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                    }
                    else if (!CmdHelper.isManualPause && !CmdHelper.isManualStop) // 运行中按空格键
                    {
                        // 暂停
                        this.BtnPause.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                    }
                }
            }
        }

        private void MenuItemOnClick(object sender, RoutedEventArgs e)
        {
            if (ParamsHelper.Permission.Available && (sender.Equals(this.MenuItemUserSettingHead)))
            {
                WindowLogin wl = new WindowLogin();
                wl.ShowDialog();
                if (!wl.isSuccessful)
                    return;
            }

            if (sender.Equals(this.MenuItemUserSettingHead))
            {
                WindowCommonSettingHead wcsh1 = new WindowCommonSettingHead(0);
                wcsh1.ShowDialog();
            }
            else if (sender.Equals(this.MenuItemModifyPwd))
            {
                WindowPwdModify wpm = new WindowPwdModify();
                wpm.ShowDialog();
            }
            else if (sender.Equals(this.MenuItemAbout))
            {
                WindowAbout wa = new WindowAbout();
                wa.ShowDialog();
            }
        }

        private void BtnOnClick(object sender, RoutedEventArgs e)
        {
            // 运行
            if (sender.Equals(this.BtnRun))
            {
                // 网络未连接
                // if (!CmdHelper.frmDAE.isCanNetSuccess)
                // {
                //     MessageBox.Show((string)this.FindResource("Prompt_Host_Disconnect_Pls_Check_Again"));
                //     return;
                // }

                // 查询仓门状态
                if (ParamsHelper.IO.DoorAvailable)
                {
                    CmdHelper.DoorQuery(false);
                    Thread.Sleep(100);
                    // 仓门打开中
                    if (!CmdHelper.frmDAE.IsDoorClosed)
                    {
                        MessageBox.Show((string)this.FindResource("Prompt_Door_Is_Opened_Please_Close_And_Run"));
                        return;
                    }
                }
                // 急停按钮状态
                if (ParamsHelper.IO.EmergencyStopAvailable)
                {
                    CmdHelper.EmergencyStopQuery(false);
                    Thread.Sleep(100);
                    // 急停按钮打开中
                    if (!CmdHelper.frmDAE.IsEmergencyStopClosed)
                    {
                        MessageBox.Show((string)this.FindResource("Prompt_EmergencyStop_Is_Opened_Please_Close_And_Run"));
                        return;
                    }
                }

                // 提示没有导入文档
                if (seqList.Count == 0)
                {
                    MessageBox.Show((string)this.FindResource("Prompt_Pls_Import_File_First"));
                }
                else
                {
                    /**
                     * 计算额外多吸体积量是否超过了量程
                     */
                    var head1RangeOver = false;
                    var head2RangeOver = false;
                    for (var headIndex = 0; headIndex < ParamsHelper.HeadList.Count; headIndex++)
                    {
                        if (ParamsHelper.HeadList[headIndex].Available)
                        {
                            // var absorbAllPercent = ParamsHelper.CommonSettingList[headIndex].AbsorbAirBeforePercent +
                            //                        ParamsHelper.CommonSettingList[headIndex].AbsorbAirAfterPercent +
                            //                        ParamsHelper.CommonSettingList[headIndex].AbsorbLiquidMoreOne2MorePercent;
                            var absorbAllPercent = ParamsHelper.CommonSettingList[headIndex].AbsorbAirBeforePercent +
                                                   ParamsHelper.CommonSettingList[headIndex].AbsorbAirAfterPercent +
                                                   ParamsHelper.CommonSettingList[headIndex].AbsorbLiquidMorePercent;
                            var headLiquidRangeMax = ConstantsUtils.LiquidRangeMaxDic[ParamsHelper.HeadList[headIndex].HeadLiquidRange];
                            var headLiquidRangeReal = Convert.ToDecimal((int)ParamsHelper.HeadList[headIndex].HeadLiquidRange);
                            // 额外多吸体积
                            var additionalVolume = absorbAllPercent * 0.01m * headLiquidRangeReal;
                            // 移液头
                            if (additionalVolume >= headLiquidRangeMax)
                            {
                                if (headIndex == 0)
                                    head1RangeOver = true;
                                else if (headIndex == 1)
                                    head2RangeOver = true;
                            }
                        }
                    }
                    if (head1RangeOver || head2RangeOver)
                    {
                        var headStr = head1RangeOver ? (string)this.FindResource("Head1") : (string)this.FindResource("Head2");
                        MessageBox.Show(headStr + " " + (string)this.FindResource("Prompt_Air_Volume_Is_Greater_Than_Liquid_Range"));
                        return;
                    }

                    // 超过量程分液中途是否更换枪头
                    var replaceTipOverRange = (bool)this.CheckBoxReplaceTipOverRange.IsChecked;

                    /**
                     * 是否有足够枪头逻辑：
                     * ①不指定枪头盒盘位：需要累积剩余枪头数目
                     * ②指定枪头盒盘位：只需计算指定盘位剩余枪头数目
                     */
                    // 运行中途是否需要置满枪头盒
                    var replaceTipboxMidway = false;
                    for (var headIndex = 0; headIndex < ParamsHelper.HeadList.Count; headIndex++)
                    {
                        if (!ParamsHelper.HeadList[headIndex].Available)
                            continue;

                        // 枪头剩余数量
                        var tipTotalCountRemain = 0;
                        // 枪头盘位
                        var tipTemplates = tipTemplateDict.Values.Where(p => p.HeadUsedIndex == headIndex);
                        // 是否指定枪头盘位
                        var isTipTemplateAssign = seqList.Exists(p => p.TipTemplateAssign);

                        // 指定枪头盒盘位
                        if (isTipTemplateAssign)
                        {
                            foreach (var template in tipTemplates)
                            {
                                tipTotalCountRemain = TipHelper.CalculateTipRemainCountByHead(template);
                                var tipTotalCountNeed = 0;
                                TipHelper.CalculateTipNeedCountByHead(seqList, headIndex, template.TipBoxTemplateIndex, ref tipTotalCountNeed, replaceTipOverRange);
                                // 是否中途需要置满枪头盒
                                if (tipTotalCountNeed > tipTotalCountRemain)
                                {
                                    replaceTipboxMidway = true;
                                    break;
                                }
                            }
                        }
                        // 不指定枪头盒盘位
                        else
                        {
                            foreach (var template in tipTemplates)
                            {
                                tipTotalCountRemain += TipHelper.CalculateTipRemainCountByHead(template);
                            }
                            // 所需枪头数目 = 每条移液信息是否取枪头 + 单条移液信息中超过量程分液是否换枪头
                            var tipTotalCountNeed = 0;
                            TipHelper.CalculateTipNeedCountByHead(seqList, headIndex, null, ref tipTotalCountNeed, replaceTipOverRange);
                            // 是否中途需要置满枪头盒
                            if (tipTotalCountNeed > tipTotalCountRemain)
                            {
                                replaceTipboxMidway = true;
                                break;
                            }
                        }
                    }

                    // 运行中途需要置满枪头盒
                    if (replaceTipboxMidway)
                    {
                        if (MessageBox.Show((string)this.FindResource("Prompt_Need_Replace_Tipbox_Midway"),
                                (string)this.FindResource("Prompt"), MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK)
                        {
                            // 运行
                            RunProcedure(replaceTipOverRange);
                        }
                    }
                    else
                        // 运行
                        RunProcedure(replaceTipOverRange);
                }
            }
            // 复位
            else if (sender.Equals(this.BtnReset))
            {
                // 网络未连接
                if (!CmdHelper.frmDAE.isCanNetSuccess)
                {
                    MessageBox.Show((string)this.FindResource("Prompt_Host_Disconnect_Pls_Check_Again"));
                    return;
                }

                SetLoadingStatus(ERunStatus.Initializing);
                BackgroundProcess.RunAsync(() => CmdHelper.InitMachine(false, false), delegate (object returnResult)
                {
                    SetLoadingStatus(ERunStatus.Stop);
                });
            }
            // 推出盘位
            else if (sender.Equals(this.BtnLaunchPlate))
            {
                // 网络未连接
                if (!CmdHelper.frmDAE.isCanNetSuccess)
                {
                    MessageBox.Show((string)this.FindResource("Prompt_Host_Disconnect_Pls_Check_Again"));
                    return;
                }

                SetLoadingStatus(ERunStatus.Initializing);
                BackgroundProcess.RunAsync(() => CmdHelper.InitMachineAndEasy2Put(false, false), delegate (object returnResult)
                {
                    SetLoadingStatus(ERunStatus.Stop);
                });
            }
            // 导入
            else if (sender.Equals(this.BtnImportData))
            {
                ShowImportDialog();
            }
            // 暂停
            else if (sender.Equals(this.BtnPause))
            {
                SetLoadingStatus(ERunStatus.Pause);
            }
            // 停止
            else if (sender.Equals(this.BtnStop))
            {
                SetLoadingStatus(ERunStatus.Stop);
                // 检查移液头状态
                CheckHeadStatus();
            }
            // 继续
            else if (sender.Equals(this.BtnContinue))
            {
                SetLoadingStatus(ERunStatus.Continue);
            }
        }

        /// <summary>
        /// 检查移液头状态
        /// </summary>
        private void CheckHeadStatus()
        {
            var head1Status = CmdHelper.headStatusList[0];
            var head2Status = CmdHelper.headStatusList[1];
            // 如果枪头有液体，提示是否把液体喷回源孔
            if (head1Status.Head == EHeadStatus.Absorbed || head2Status.Head == EHeadStatus.Absorbed || head1Status.VolumeAbsorbMoreLeft > 0 || head2Status.VolumeAbsorbMoreLeft > 0)
            {
                // 哪个移液头
                var headUsedIndex = head1Status.Head == EHeadStatus.Absorbed || head1Status.VolumeAbsorbMoreLeft > 0 ? 0 : 1;
                var seq = seqList[CmdHelper.headStatusList[headUsedIndex].SeqIndex];
                if (ParamsHelper.IO.SerialPortAvailable || MessageBox.Show((string)this.FindResource("Prompt_If_Jet_Liquid_2_Source_Hole"),
                        (string)this.FindResource("Prompt"), MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                {
                    var headStatus = CmdHelper.headStatusList[headUsedIndex];
                    if (headStatus.Head == EHeadStatus.Absorbed || headStatus.VolumeAbsorbMoreLeft > 0)
                    {
                        BackgroundProcess.RunAsync(() => CmdHelper.ReJet2Source(headUsedIndex, seq, headStatus.SourceConsumableType, headStatus.SourceTemplateIndex, headStatus.SourceHoleIndex, false), delegate (object returnResult)
                        {
                            // 是否退枪头
                            PromptReleaseTip(headUsedIndex, false);
                        });
                    }
                }
                else
                {
                    // 是否退枪头
                    PromptReleaseTip(headUsedIndex, false);
                }
            }
            // 如果取枪头了，提示是否到退枪头位退枪头
            else if (head1Status.Head == EHeadStatus.TipTook || head2Status.Head == EHeadStatus.TipTook)
            {
                // 哪个移液头
                var headUsedIndex = head1Status.Head == EHeadStatus.TipTook ? 0 : 1;
                PromptReleaseTip(headUsedIndex, false);
            }
            // 如果喷液了，提示是否到退枪头位退枪头
            else if (head1Status.Head == EHeadStatus.Jetted || head2Status.Head == EHeadStatus.Jetted)
            {
                // 哪个移液头
                var headUsedIndex = head1Status.Head == EHeadStatus.Jetted ? 0 : 1;
                PromptReleaseTip(headUsedIndex, false);
            }
        }

        /// <summary>
        /// 提示是否退枪头
        /// </summary>
        /// <param name="headUsedIndex"></param>
        /// <param name="isNeedZ0AfterRelease">退枪头后是否需要高度复位</param>
        private void PromptReleaseTip(int headUsedIndex, bool isNeedZ0AfterRelease)
        {
            // 是否退枪头
            if (ParamsHelper.IO.SerialPortAvailable || MessageBox.Show((string)this.FindResource("Prompt_If_Release_Tip"),
                    (string)this.FindResource("Prompt"), MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                new Thread(() =>
                {
                    try
                    {
                        ActionReleaseTip(headUsedIndex, isNeedZ0AfterRelease, true, 0, false);
                    }
                    catch (ManualStopException e)
                    {
                    }
                }).Start();
            }
        }

        /// <summary>
        /// 执行程序（带耗材摆放提示）
        /// </summary>
        /// <param name="replaceTipOverRange">超过量程分液是否更换枪头</param>
        private void RunProcedure(bool replaceTipOverRange)
        {
            // 提示是否已经摆放好板
            if (MessageBox.Show((string)this.FindResource("Prompt_Need_To_Put_Template_OK"),
                    (string)this.FindResource("Prompt"), MessageBoxButton.OKCancel, MessageBoxImage.Warning) ==
                MessageBoxResult.OK)
            {
                RunProcedure(replaceTipOverRange, true);
            }
        }

        /// <summary>
        /// 执行程序
        /// </summary>
        /// <param name="replaceTipOverRange">超过量程分液是否更换枪头</param>
        /// <param name="isSuccessHint">完成后是否弹出框提示</param>
        /// <param name="completedCallback">执行完成回调（与第三方设备通信使用）</param>
        private void RunProcedure(bool replaceTipOverRange, bool isSuccessHint, MyCallBack completedCallback = null)
        {
            Dispatcher.Invoke(() => { SetLoadingStatus(ERunStatus.Running); });
            var thread = new Thread(() =>
            {
                try
                {
                    LogHelper.InfoTitle((string)this.FindResource("ExecFile") + "《" + excelFilesNameImported + "》");

                    /**
                     * 计算多点校准
                     */
                    CalcCalibration();

                    // 运行前发送速度指令
                    for (var headIndex = 0; headIndex < ParamsHelper.HeadList.Count; headIndex++)
                    {
                        if (ParamsHelper.HeadList[headIndex].Available)
                            CmdHelper.SpeedSet(headIndex, EAxis.All, "", 100, true);
                    }

                    // 初始化
                    CmdHelper.InitMachine(true, false);

                    TakeAction(replaceTipOverRange);

                    // 初始化
                    CmdHelper.InitMachineAndEasy2Put(true, false);

                    LogHelper.InfoTail((string)this.FindResource("ExecCompleted"));

                    Dispatcher.Invoke(() =>
                    {
                        SetLoadingStatus(ERunStatus.Stop);
                        // 如果枪头退回到取枪头位置，就把取枪头位置自动置满
                        for (var headIndex = 0; headIndex < ParamsHelper.CommonSettingList.Count; headIndex++)
                        {
                            var commonSetting = ParamsHelper.CommonSettingList[headIndex];
                            if (commonSetting.ReleaseTipBack2TakePos)
                            {
                                var tipTemplates = tipTemplateDict.Values.Where(p => p.HeadUsedIndex == headIndex);
                                foreach (var tipTemplate in tipTemplates)
                                {
                                    tipTemplate.SplitButtonTipsBoxPos.SelectedIndex = 0;
                                }
                            }
                        }
                        if (isSuccessHint)
                            MessageBox.Show((string)this.FindResource("Prompt_Liquid_Relief_Success"));
                    });

                    if (completedCallback != null)
                        completedCallback.Callback();
                }
                catch (ManualStopException)
                {
                }
            });
            thread.IsBackground = true;
            thread.Start();
        }

        /// <summary>
        /// 导入文档
        /// </summary>
        private void ShowImportDialog()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            var result = dialog.ShowDialog();
            if (result == true)
            {
                // 导入excel
                excelFilesNameImported = dialog.SafeFileName;
                ImportExcel(dialog.FileName);
            }
        }

        /// <summary>
        /// 导入Excel
        /// </summary>
        /// <param name="filePath">文件路径</param>
        private void ImportExcel(string filePath)
        {
            // 解析excel数据
            this.LayoutImportLoading.Visibility = Visibility.Visible;
            BackgroundProcess.RunAsync(() => ParseDataFromExcel(filePath), delegate (object returnResult)
            {
                this.LayoutImportLoading.Visibility = Visibility.Collapsed;
                if ((bool)returnResult)
                {
                    MessageBox.Show((string)this.FindResource("Prompt_Import_Success"));
                }
                else
                {
                    MessageBox.Show((string)this.FindResource("Prompt_Import_Error"));
                }
            });
        }

        /// <summary>
        /// 解析Excel数据
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns></returns>
        private bool ParseDataFromExcel(string filePath)
        {
            using (Workbook workbook = new Workbook())
            {
                try
                {
                    workbook.LoadFromFile(filePath);
                }
                catch (Exception e)
                {
                    MessageBox.Show((string)this.FindResource("Prompt_Pls_Close_Excel_File_First"));
                    LogHelper.Error(e.StackTrace);
                    return false;
                }

                Worksheet ws = null;
                try
                {
                    ws = workbook.Worksheets[0];
                }
                catch (Exception ex)
                {
                    MessageBox.Show((string)this.FindResource("Prompt_Excel_Must_Be_Xlsx_File_Or_File_Destroy"));
                    LogHelper.Error(ex.StackTrace);
                    return false;
                }

                // 先清空原数据
                seqList.Clear();
                // 读取的开始行
                var rowIndex = 3;
                try
                {
                    /**
                     * Excel列
                     */
                    const int tipTemplateIndexCol = 1; // 枪头盘位列
                    const int tipTemplateConsumableTypeCol = 2;  // 枪头盘耗材类型列

                    const int sourceTemplateIndexCol = 3; // 源盘盘位列
                    const int sourceTemplateConsumableTypeCol = 4;  // 源盘耗材类型列
                    const int sourceHoleIndexCol = 5; // 源盘孔位置列

                    const int targetTemplateIndexCol = 7; // 靶盘盘位列
                    const int targetTemplateConsumableTypeCol = 8;  // 靶盘耗材类型列
                    const int targetHoleIndexCol = 9; // 靶盘孔位置列
                    const int volumeCol = 10; // 每孔体积列

                    const int absorbMixingVolumeCol = 11; // 吸前混合体积
                    const int absorbMixingCountCol = 12; // 吸前混合次数
                    const int jetMixingVolumeCol = 13; // 喷后混合体积
                    const int jetMixingCountCol = 14; // 喷后混合次数

                    const int absorbWallCol = 15;  // 吸液靠壁
                    const int jetWallCol = 16;  // 喷液靠壁

                    const int replaceTipCol = 17; // 是否插枪头
                    const int tipChannelCol = 18; // 取枪头数目

                    int cmdCol = 18; // 特殊指令
                    // 兼容Excel包含“取枪头数目”版本
                    var isNeedTipChannel = false;
                    if (null != ws.Range[1, 19].Text && ws.Range[1, 19].Text.Trim().Contains((string)this.FindResource("SpecialCmd")))
                    {
                        cmdCol = 19;
                        isNeedTipChannel = true;
                    }

                    // 判断是否存在数据
                    if (null == ws.Range[rowIndex, sourceTemplateIndexCol].Value && null == ws.Range[rowIndex, cmdCol].Text)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            MessageBox.Show(
                                (string)this.FindResource(
                                    "Prompt_Import_Excel_Must_Not_Be_Empty"));
                        });
                        return false;
                    }

                    // 逐一解析每行数据
                    while (true)
                    {
                        // 源盘 、靶盘板位 和 特殊指令
                        var sourceTemplateIndexEmpty = ViewUtils.CheckExcelCellEmpty(ws.Range[rowIndex, sourceTemplateIndexCol]);
                        var targetTemplateIndexEmpty = ViewUtils.CheckExcelCellEmpty(ws.Range[rowIndex, targetTemplateIndexCol]);
                        var cmdContent = ws.Range[rowIndex, cmdCol].Text;
                        var cmdEmpty = ViewUtils.CheckExcelCellEmpty(ws.Range[rowIndex, cmdCol]);
                        // 注意：targetTemplateIndexEmpty为null 不能简单判断，还需要判断内容是否为空
                        if (!targetTemplateIndexEmpty || !cmdEmpty)
                        {
                            // 特殊指令
                            var cmd = cmdEmpty ? "" : cmdContent.Trim();

                            // txt文件链接（执行txt内容）
                            var isTxtLink = false;
                            if (cmd.ToLower().Contains(".txt"))
                            {
                                string rootPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                                var txtFilesNameList = Directory.EnumerateFiles(rootPath, "*.txt", SearchOption.TopDirectoryOnly)
                                    .Select(f => Path.GetFileName(f)).Where(p => !p.StartsWith("AutoLiquid_Update_Log") && p.StartsWith(cmd.ToLower())).ToList();
                                var txtFilesNameListCount = txtFilesNameList.Count();
                                if (txtFilesNameListCount != 0)
                                    cmd = File.ReadAllText(rootPath + @"\" + txtFilesNameList.ElementAt(0)).Trim();
                                else
                                {
                                    Dispatcher.Invoke(() =>
                                    {
                                        MessageBox.Show(
                                            (string)this.FindResource(
                                                "Prompt_Txt_Link_File_Not_Exist"));
                                    });
                                    return false;
                                }

                                isTxtLink = true;
                            }

                            // 吸液、喷液前后特殊指令
                            var cmdAbsorbBefore = ObjectUtils.GetCmdAccordTag(cmd,
                                AutoLiquid_Library.Utils.ConstantsUtils.AbsorbBeforeCmd, isTxtLink);
                            var cmdAbsorbAfter = ObjectUtils.GetCmdAccordTag(cmd,
                                AutoLiquid_Library.Utils.ConstantsUtils.AbsorbAfterCmd, isTxtLink);
                            var cmdJetBefore = ObjectUtils.GetCmdAccordTag(cmd,
                                AutoLiquid_Library.Utils.ConstantsUtils.JetBeforeCmd, isTxtLink);
                            var cmdJetAfter = ObjectUtils.GetCmdAccordTag(cmd,
                                AutoLiquid_Library.Utils.ConstantsUtils.JetAfterCmd, isTxtLink);

                            // 吸液混合、喷液混合前特殊指令
                            var cmdAbsorbMixingBefore = ObjectUtils.GetCmdAccordTag(cmd,
                                AutoLiquid_Library.Utils.ConstantsUtils.AbsorbMixingBeforeCmd, isTxtLink);
                            var cmdJetMixingBefore = ObjectUtils.GetCmdAccordTag(cmd,
                                AutoLiquid_Library.Utils.ConstantsUtils.JetMixingBeforeCmd, isTxtLink);
                            // 吸液混合、喷液混合后特殊指令
                            var cmdAbsorbMixingAfter = ObjectUtils.GetCmdAccordTag(cmd,
                                AutoLiquid_Library.Utils.ConstantsUtils.AbsorbMixingAfterCmd, isTxtLink);
                            var cmdJetMixingAfter = ObjectUtils.GetCmdAccordTag(cmd,
                                AutoLiquid_Library.Utils.ConstantsUtils.JetMixingAfterCmd, isTxtLink);

                            // 是否只含有特殊指令
                            if (targetTemplateIndexEmpty)
                            {
                                seqList.Add(new Seq { IsCmdOnly = true, Cmd = cmd });
                                rowIndex++;
                                continue;
                            }

                            // 使用的移液头Index
                            var headUsedIndex = cmd.ToLower().Contains(AutoLiquid_Library.Utils.ConstantsUtils.Head2Cmd.ToLower()) && ParamsHelper.HeadList[1].Available ? 1 : 0;

                            /**
                             * 靶盘
                             */
                            // 靶盘盘位Index
                            var targetTemplateIndexList = ws.Range[rowIndex, targetTemplateIndexCol].Value.Trim()
                                .Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(x => int.Parse(x) - 1).ToList();
                            // 靶盘耗材类型
                            var targetTemplateConsumableType = ConsumableHelper.GetConsumableType(headUsedIndex, ws.Range[rowIndex, targetTemplateConsumableTypeCol].Text.Trim(), false);
                            // 检查是否存在该耗材名字
                            if (targetTemplateConsumableType == null)
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    MessageBox.Show(
                                        (string)this.FindResource(
                                            "Prompt_Import_Excel_Group_Name_Not_Exist"));
                                });
                                return false;
                            }
                            // 检查是否已经启用该盘位
                            foreach (var targetTemplateIndex in targetTemplateIndexList)
                            {
                                if (!targetTemplateConsumableType.TemplateAvailableList[targetTemplateIndex])
                                {
                                    Dispatcher.Invoke(() =>
                                    {
                                        var prompt = (string)this.FindResource(
                                                "Prompt_Pls_Enable_Template_In_User_Setting_1") + targetTemplateConsumableType
                                                                                                    .GroupName
                                                                                                + (string)this
                                                                                                    .FindResource(
                                                                                                        "Prompt_Pls_Enable_Template_In_User_Setting_2") + (targetTemplateIndex + 1);
                                        MessageBox.Show(prompt);
                                    });
                                    return false;
                                }
                            }
                            // 靶盘位置孔Index
                            var targetHoleIndexList = ws.Range[rowIndex, targetHoleIndexCol].Value.Trim()
                                .Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(x => ConsumableHelper.GetHoleIndex(headUsedIndex, targetTemplateConsumableType, x.Trim())).ToList();
                            // 体积
                            var volumeList = ws.Range[rowIndex, volumeCol].Value.Trim()
                                .Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(x => decimal.Parse(x)).Select(d => new Volume { Original = d }).ToList();

                            // 是否泵分液
                            if (sourceTemplateIndexEmpty && !targetTemplateIndexEmpty && !cmdEmpty)
                            {
                                seqList.Add(new Seq
                                {
                                    TargetTemplateIndexList = targetTemplateIndexList,
                                    TargetTemplateConsumableType = targetTemplateConsumableType,
                                    TargetHoleIndexList = targetHoleIndexList,
                                    VolumeEachList = volumeList,
                                    Cmd = cmd,
                                    IsPumpLiquid = true
                                });
                                rowIndex++;
                                continue;
                            }

                            /**
                             * 枪头
                             */
                            // 枪头盘位Index
                            var tipTemplateIndex = short.Parse(ws.Range[rowIndex, tipTemplateIndexCol].Value) - 1;
                            // 是否指定盘位取枪头
                            var tipTemplateAssign = ws.Range["A1"].Text.Trim().Contains((string)this.FindResource("TemplateAssign"));
                            // 枪头盒耗材类型
                            var tipTemplateConsumableType = ConsumableHelper.GetConsumableType(headUsedIndex, ws.Range[rowIndex, tipTemplateConsumableTypeCol].Text.Trim(), true);
                            // 检查是否存在该耗材名字
                            if (tipTemplateConsumableType == null)
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    MessageBox.Show(
                                        (string)this.FindResource(
                                            "Prompt_Import_Excel_Group_Name_Not_Exist"));
                                });
                                return false;
                            }
                            // 检查是否已经启用该盘位
                            if (!tipTemplateConsumableType.TemplateAvailableList[tipTemplateIndex])
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    var prompt = (string)this.FindResource(
                                            "Prompt_Pls_Enable_Template_In_User_Setting_1") + tipTemplateConsumableType
                                                                                                .GroupName
                                                                                            + (string)this
                                                                                                .FindResource(
                                                                                                    "Prompt_Pls_Enable_Template_In_User_Setting_2") + (tipTemplateIndex + 1);
                                    MessageBox.Show(prompt);
                                });
                                return false;
                            }

                            /**
                             * 源盘
                             */
                            // 源盘盘位Index
                            var sourceTemplateIndex = short.Parse(ws.Range[rowIndex, sourceTemplateIndexCol].Value) - 1;
                            // 源盘耗材类型
                            var sourceTemplateConsumableType = ConsumableHelper.GetConsumableType(headUsedIndex, ws.Range[rowIndex, sourceTemplateConsumableTypeCol].Text.Trim(), false);
                            // 检查是否存在该耗材名字
                            if (sourceTemplateConsumableType == null)
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    MessageBox.Show(
                                        (string)this.FindResource(
                                            "Prompt_Import_Excel_Group_Name_Not_Exist"));
                                });
                                return false;
                            }
                            // 检查是否已经启用该盘位
                            if (!sourceTemplateConsumableType.TemplateAvailableList[sourceTemplateIndex])
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    var prompt = (string)this.FindResource(
                                            "Prompt_Pls_Enable_Template_In_User_Setting_1") + sourceTemplateConsumableType
                                                                                                .GroupName
                                                                                            + (string)this
                                                                                                .FindResource(
                                                                                                    "Prompt_Pls_Enable_Template_In_User_Setting_2") + (sourceTemplateIndex + 1);
                                    MessageBox.Show(prompt);
                                });
                                return false;
                            }
                            // 源盘位置孔Index
                            // var sourceHoleIndex =
                            //     ConsumableHelper.GetHoleIndex(headUsedIndex, sourceTemplateConsumableType, ws.Range[rowIndex, sourceHoleIndexCol].Text.Trim());
                            var sourceHoleIndexList = ws.Range[rowIndex, sourceHoleIndexCol].Value.Trim()
                                .Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(x => ConsumableHelper.GetHoleIndex(headUsedIndex, sourceTemplateConsumableType, x.Trim())).ToList();


                            /**
                             * 混合信息
                             */
                            // 吸前混合
                            var absorbMixingVolumeContent = ws.Range[rowIndex, absorbMixingVolumeCol].Value;
                            var absorbMixingVolume = null == absorbMixingVolumeContent || absorbMixingVolumeContent.Equals("") ? 0m : decimal.Parse(absorbMixingVolumeContent.Trim());
                            var absorbMixingCountContent = ws.Range[rowIndex, absorbMixingCountCol].Value;
                            var absorbMixingCount = null == absorbMixingCountContent || absorbMixingCountContent.Equals("") ? 0 : short.Parse(absorbMixingCountContent.Trim());
                            // 喷后混合
                            var jetMixingVolumeContent = ws.Range[rowIndex, jetMixingVolumeCol].Value;
                            var jetMixingVolume = null == jetMixingVolumeContent || jetMixingVolumeContent.Equals("") ? 0m : decimal.Parse(jetMixingVolumeContent.Trim());
                            var jetMixingCountContent = ws.Range[rowIndex, jetMixingCountCol].Value;
                            var jetMixingCount = null == jetMixingCountContent || jetMixingCountContent.Equals("") ? 0 : short.Parse(jetMixingCountContent.Trim());

                            /**
                             * 靠壁
                             */
                            // 吸液靠壁
                            var absorbWall = null == ws.Range[rowIndex, absorbWallCol].Text ? "" : ws.Range[rowIndex, absorbWallCol].Text.Trim();
                            var absorbWallList = ExcelHelper.GetWallList(absorbWall);
                            // 喷液靠壁
                            var jetWall = null == ws.Range[rowIndex, jetWallCol].Text ? "" : ws.Range[rowIndex, jetWallCol].Text.Trim();
                            var jetWallList = ExcelHelper.GetWallList(jetWall);

                            // 是否注释该行序列
                            var isComment = cmd.Contains("//");

                            // 是否取枪头
                            var isTakeTip = ws.Range[rowIndex, replaceTipCol].Text.Trim().Equals((string)this.FindResource("Yes"));
                            // 取枪头数目
                            var tipChannel = isNeedTipChannel && !ViewUtils.CheckExcelCellEmpty(ws.Range[rowIndex, tipChannelCol]) ? TipHelper.GetTipChannel2DArray(headUsedIndex, int.Parse(ws.Range[rowIndex, tipChannelCol].Value.Trim()))
                                : new int[ParamsHelper.HeadList[headUsedIndex].ChannelRow, ParamsHelper.HeadList[headUsedIndex].ChannelCol];

                            // 是否用于梯度稀释
                            var serialDilute = targetTemplateConsumableType.JetMixingHeight > 0 && jetMixingVolume > 0 && jetMixingCount > 0 && targetHoleIndexList.Count > 1;

                            // 喷液体积补偿
                            var cmdJetOffset = ObjectUtils.GetCmdAccordTag(cmd, AutoLiquid_Library.Utils.ConstantsUtils.JetOffsetCmd, isTxtLink);
                            var jetOffsetList = GetJetOffset(cmdJetOffset);

                            // 一吸多喷喷液后回吸体积
                            var cmdVolumeBackAbsorb = ObjectUtils.GetCmdAccordTag(cmd, AutoLiquid_Library.Utils.ConstantsUtils.BackAbsorbCmd, isTxtLink);
                            var volumeBackAbsorb = cmdVolumeBackAbsorb.Equals("") ? 0 : decimal.Parse(cmdVolumeBackAbsorb);

                            // 吸液后多吸体积
                            var cmdSourceVolumeAbsorbMore = ObjectUtils.GetCmdAccordTag(cmd, AutoLiquid_Library.Utils.ConstantsUtils.AbsorbMoreCmd, isTxtLink);
                            var sourceVolumeAbsorbMore = cmdSourceVolumeAbsorbMore.Equals("") ? 0 : decimal.Parse(cmdSourceVolumeAbsorbMore);

                            // 吸后反喷体积
                            var cmdSourceVolumeReverseJet = ObjectUtils.GetCmdAccordTag(cmd, AutoLiquid_Library.Utils.ConstantsUtils.ReverseJetCmd, isTxtLink);
                            var sourceVolumeReverseJet = cmdSourceVolumeReverseJet.Equals("") ? 0 : decimal.Parse(cmdSourceVolumeReverseJet);

                            // 多吸液体返回源孔喷出
                            var reJet2Source = cmd.ToLower().Contains(AutoLiquid_Library.Utils.ConstantsUtils.ReJet2SourceCmd.ToLower());

                            var seq = new Seq
                            {
                                TipTemplateIndex = tipTemplateIndex,
                                TipTemplateAssign = tipTemplateAssign,
                                TipTemplateConsumableType = tipTemplateConsumableType,
                                IsTakeTip = isTakeTip,
                                TipChannel = tipChannel,
                                SourceTemplateIndex = sourceTemplateIndex,
                                SourceTemplateConsumableType = sourceTemplateConsumableType,
                                // SourceHoleIndex = sourceHoleIndex,
                                SourceHoleIndexList = sourceHoleIndexList, // 20250625 SourceHoleIndex改成SourceHoleIndexList
                                SourceVolumeAbsorbMore = sourceVolumeAbsorbMore,
                                SourceVolumeReverseJet = sourceVolumeReverseJet,
                                ReJet2Source = reJet2Source,
                                TargetTemplateIndexList = targetTemplateIndexList,
                                TargetTemplateConsumableType = targetTemplateConsumableType,
                                TargetHoleIndexList = targetHoleIndexList,
                                SerialDilute = serialDilute,
                                JetOffsetList = jetOffsetList,
                                VolumeBackAbsorb = volumeBackAbsorb,
                                VolumeEachList = volumeList,
                                AbsorbMixingVolume = absorbMixingVolume,
                                AbsorbMixingCount = absorbMixingCount,
                                JetMixingVolume = jetMixingVolume,
                                JetMixingCount = jetMixingCount,
                                AbsorbWallList = absorbWallList,
                                JetWallList = jetWallList,
                                IsTxtLink = isTxtLink,
                                Cmd = cmd,
                                CmdAbsorbBefore = cmdAbsorbBefore,
                                CmdAbsorbAfter = cmdAbsorbAfter,
                                CmdJetBefore = cmdJetBefore,
                                CmdJetAfter = cmdJetAfter,
                                CmdAbsorbMixingBefore = cmdAbsorbMixingBefore,
                                CmdAbsorbMixingAfter = cmdAbsorbMixingAfter,
                                CmdJetMixingBefore = cmdJetMixingBefore,
                                CmdJetMixingAfter = cmdJetMixingAfter,
                                IsComment = isComment,
                                HeadUsedIndex = headUsedIndex
                            };
                            seqList.Add(seq);
                        }
                        else
                            break;

                        rowIndex++;
                    }


                    try
                    {
                        // 初始化盘位信息
                        Dispatcher.Invoke(InitTemplates);
                        // 重新计算吸液、喷液位置（仅适用于灵活取枪头，如不需要灵活取枪头，可不调用）
                        AllocateAbsorbAndJetInfo();
                        // 分配退枪头信息
                        AllocateReleaseTipInfo();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show((string)this.FindResource("Prompt_Pls_Check_Excel_And_Consumable_Info"));
                        LogHelper.Error(ex.StackTrace);
                        return false;
                    }

                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show((string)this.FindResource("Prompt_Excel_Error_1") + rowIndex +
                                    (string)this.FindResource("Prompt_Excel_Error_2") + "：" + "");
                    LogHelper.Error(ex.StackTrace);
                    return false;
                }
            }
        }

        /// <summary>
        /// 获取喷液补偿信息
        /// </summary>
        /// <param name="cmdJetOffset"></param>
        /// <returns></returns>
        private List<JetOffset> GetJetOffset(string cmdJetOffset)
        {
            var jetOffsetList = new List<JetOffset>();
            if (!cmdJetOffset.Equals(""))
            {
                var items = cmdJetOffset.Split(new[] { "(", ")" }, StringSplitOptions.RemoveEmptyEntries).Where(p => !p.Equals(",")).ToList();
                foreach (var item in items)
                {
                    var content = item.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries);
                    var posStr = content[0];
                    var volume = decimal.Parse(content[1]);
                    // 判断是否孔位范围
                    if (posStr.Contains("-") || posStr.Contains("~"))
                    {
                        var posRange = posStr.Split(new[] { '-', '~' }, StringSplitOptions.RemoveEmptyEntries);
                        var fromPosIndex = int.Parse(posRange[0]) - 1;
                        var toPosIndex = int.Parse(posRange[1]) - 1;
                        for (var index = fromPosIndex; index <= toPosIndex; index++)
                        {
                            jetOffsetList.Add(new JetOffset { PosIndex = index, VolumeOffset = volume });
                        }
                    }
                    else
                    {
                        var posIndex = int.Parse(posStr) - 1;
                        var jetOffset = jetOffsetList.FirstOrDefault(p => p.PosIndex == posIndex);
                        if (null != jetOffset)
                            jetOffset.VolumeOffset = volume;
                        else
                            jetOffsetList.Add(new JetOffset { PosIndex = posIndex, VolumeOffset = volume });
                    }
                }
            }

            return jetOffsetList;
        }

        /// <summary>
        /// 计算多点校准
        /// </summary>
        private void CalcCalibration()
        {
            foreach (var seq in seqList)
            {
                var headIndex = seq.HeadUsedIndex;
                foreach (var volume in seq.VolumeEachList)
                {
                    volume.Calibration = Utils.DataHelper.CalibrateVol(headIndex, volume.Original);
                }
            }
        }

        /// <summary>
        /// 初始化盘位信息
        /// </summary>
        private void InitTemplates()
        {
            // 先清空原盘位数据
            foreach (var t in templateCanUse)
            {
                t.Children.Clear();
            }

            tipTemplateDict.Clear();
            sourceTemplateDict.Clear();
            targetTemplateDict.Clear();

            // 由于可能存在ep管架之类的占据多个盘面的耗材，所以先复位每个盘位占据盘面个数为1
            for (var i = 0; i < this.GridTemplateContainer.Children.Count; i++)
            {
                var controlTemplate = this.GridTemplateContainer.Children[i] as UserControls.ControlTemplate;
                Grid.SetRow(controlTemplate, 0);
                Grid.SetRowSpan(controlTemplate, ConstantsUtils.TemplateOccupyGridSpan);
                Grid.SetColumn(controlTemplate, 0);
                Grid.SetColumnSpan(controlTemplate, ConstantsUtils.TemplateOccupyGridSpan);
            }

            // 枪头盒、源盘、靶盘盘位标题计数器
            var tipTemplateTitleTick = 0;
            var sourceTemplateTitleTick = 0;
            var targetTemplateTitleTick = 0;
            foreach (var seq in seqList)
            {
                // 只含有特殊指令
                if (seq.IsCmdOnly)
                    continue;

                var tipTemplateConsumableType = seq.TipTemplateConsumableType;
                var tipTemplateIndex = seq.TipTemplateIndex;
                var tipChannel = seq.TipChannel;
                var sourceTemplateConsumableType = seq.SourceTemplateConsumableType;
                var sourceTemplateIndex = seq.SourceTemplateIndex;
                // var sourceHoleIndex = seq.SourceHoleIndex;
                var sourceHoleIndexList = seq.SourceHoleIndexList; // 20250625 SourceHoleIndex改成SourceHoleIndexList
                var volumeEachList = seq.VolumeEachList;
                var targetTemplateConsumableType = seq.TargetTemplateConsumableType;
                var targetTemplateIndexList = seq.TargetTemplateIndexList;
                var targetHoleIndexList = seq.TargetHoleIndexList;
                var serialDilute = seq.SerialDilute;
                var headUsedIndex = seq.HeadUsedIndex;
                var isPumpLiquid = seq.IsPumpLiquid;


                // 枪头盘位
                if (!isPumpLiquid)
                {
                    if (!tipTemplateDict.ContainsKey(tipTemplateIndex))
                    {
                        // 设置占用盘位
                        ViewUtils.SetTemplateOccupy(headUsedIndex, this.GridTemplateContainer, tipTemplateIndex, tipTemplateConsumableType.TemplateOccupySpan);

                        tipTemplateTitleTick++;
                        var template = new Template
                        {
                            Title = (string)this.FindResource("TemplateTips") + tipTemplateTitleTick,
                            RowCount = tipTemplateConsumableType.RowCount,
                            ColCount = tipTemplateConsumableType.ColCount,
                            Step = tipTemplateConsumableType.HoleStep,
                            A1Pos = ParamsHelper.CommonSettingList[0].A1Pos,
                            Type = ETemplateType.Tip
                        };
                        var holeTotalCount = template.RowCount * template.ColCount; // 孔总数
                        for (var i = 0; i < holeTotalCount; i++)
                        {
                            template.Holes.Add(new Hole { Index = i });
                        }
                        // 枪头盒是否灵活取枪头（判断所有seq是否含有自定义取枪头数目，如果是，则认为是灵活取枪头）
                        var tipBoxFlexible = TipHelper.IsTipBoxFlexible(headUsedIndex, seqList, tipTemplateConsumableType);
                        var controlTemplateTip = new ControlTemplateTip(template, tipTemplateConsumableType) { TipBoxTemplateIndex = tipTemplateIndex, TipBoxFlexible = tipBoxFlexible, HeadUsedIndex = headUsedIndex };
                        tipTemplateDict.Add(tipTemplateIndex, controlTemplateTip);
                        templateCanUse[tipTemplateIndex].Children.Add(controlTemplateTip);
                    }
                }

                // 源盘盘位
                if (!isPumpLiquid)
                {
                    var sourceTotalVolume = 0m; // 源孔所需液体量
                    for (var i = 0; i < targetHoleIndexList.Count; i++)
                    {
                        if (serialDilute)
                            sourceTotalVolume = volumeEachList[0].Original;
                        // 如果一吸多喷投入体积只有一个，就默认为相同体积，否则体积数必须等于靶孔数
                        else if (volumeEachList.Count > 1)
                            sourceTotalVolume += volumeEachList[i].Original;
                        else
                            sourceTotalVolume += volumeEachList[0].Original;
                    }
                    if (!sourceTemplateDict.ContainsKey(sourceTemplateIndex))
                    {
                        // 设置占用盘位
                        ViewUtils.SetTemplateOccupy(headUsedIndex, this.GridTemplateContainer, sourceTemplateIndex, sourceTemplateConsumableType.TemplateOccupySpan);

                        sourceTemplateTitleTick++;
                        var template = new Template
                        {
                            Title = (string)this.FindResource("TemplateSource") + sourceTemplateTitleTick,
                            SubTitle = sourceTemplateConsumableType.GroupName,
                            RowCount = sourceTemplateConsumableType.RowCount,
                            ColCount = sourceTemplateConsumableType.ColCount,
                            Step = sourceTemplateConsumableType.HoleStep,
                            A1Pos = ParamsHelper.CommonSettingList[0].A1Pos,
                            Type = ETemplateType.Source
                        };
                        var holeTotalCount = template.RowCount * template.ColCount; // 孔总数
                        for (var i = 0; i < holeTotalCount; i++)
                        {
                            template.Holes.Add(new Hole { Index = i, Capacity = 0 });
                        }

                        var controlTemplate = new ControlTemplateCommon(template, sourceTemplateConsumableType);
                        // controlTemplate.RefreshTemplateHolesStatus(headUsedIndex, sourceHoleIndex, tipChannel, sourceTotalVolume);
                        // 20250625 SourceHoleIndex改成SourceHoleIndexList
                        foreach (var sourceHoleIndex in sourceHoleIndexList)
                        {
                            controlTemplate.RefreshTemplateHolesStatus(headUsedIndex, sourceHoleIndex, tipChannel, sourceTotalVolume);
                        }
                        sourceTemplateDict.Add(sourceTemplateIndex, controlTemplate);
                        templateCanUse[sourceTemplateIndex].Children.Add(controlTemplate);
                    }
                    else
                    {
                        if (!sourceTemplateDict.TryGetValue(sourceTemplateIndex, out var controlTemplate))
                            return; // 该源盘位Index不存在
                        // controlTemplate.RefreshTemplateHolesStatus(headUsedIndex, sourceHoleIndex, tipChannel, sourceTotalVolume);
                        // 20250625 SourceHoleIndex改成SourceHoleIndexList
                        foreach (var sourceHoleIndex in sourceHoleIndexList)
                        {
                            controlTemplate.RefreshTemplateHolesStatus(headUsedIndex, sourceHoleIndex, tipChannel, sourceTotalVolume);
                        }
                    }
                }

                // 靶盘盘位
                var targetTemplateIndexCount = targetTemplateIndexList.Count;
                for (var i = 0; i < targetTemplateIndexCount; i++)
                {
                    var targetTemplateIndex = targetTemplateIndexList[i];

                    if (!targetTemplateDict.ContainsKey(targetTemplateIndex))
                    {
                        // 设置占用盘位
                        ViewUtils.SetTemplateOccupy(headUsedIndex, this.GridTemplateContainer, targetTemplateIndex, targetTemplateConsumableType.TemplateOccupySpan);

                        targetTemplateTitleTick++;
                        var template = new Template
                        {
                            Title = (string)this.FindResource("TemplateTarget") + targetTemplateTitleTick,
                            SubTitle = targetTemplateConsumableType.GroupName,
                            RowCount = targetTemplateConsumableType.RowCount,
                            ColCount = targetTemplateConsumableType.ColCount,
                            Step = targetTemplateConsumableType.HoleStep,
                            A1Pos = ParamsHelper.CommonSettingList[0].A1Pos,
                            Type = ETemplateType.Target
                        };
                        var holeTotalCount = template.RowCount * template.ColCount; // 孔总数
                        for (var j = 0; j < holeTotalCount; j++)
                        {
                            template.Holes.Add(new Hole { Index = j, Capacity = 0 });
                        }

                        var controlTemplate = new ControlTemplateCommon(template, targetTemplateConsumableType);

                        // 多靶多孔
                        if (targetTemplateIndexCount > 1)
                        {
                            var targetVolumeEach = volumeEachList.Count > 1 ? volumeEachList[i] : volumeEachList[0];
                            var targetVolumeTotal = targetVolumeEach.Original * sourceHoleIndexList.Count;
                            controlTemplate.RefreshTemplateHolesStatus(headUsedIndex, targetHoleIndexList[i], tipChannel, targetVolumeTotal);
                        }
                        // 1靶多孔
                        else
                        {
                            for (var j = 0; j < targetHoleIndexList.Count; j++)
                            {
                                var targetHoleIndex = targetHoleIndexList[j];
                                var targetVolumeEach = volumeEachList.Count > 1 ? volumeEachList[j] : volumeEachList[0];
                                var targetVolumeTotal = targetVolumeEach.Original * sourceHoleIndexList.Count;
                                controlTemplate.RefreshTemplateHolesStatus(headUsedIndex, targetHoleIndex, tipChannel, targetVolumeTotal);
                            }
                        }

                        targetTemplateDict.Add(targetTemplateIndex, controlTemplate);
                        templateCanUse[targetTemplateIndex].Children.Add(controlTemplate);
                    }
                    else
                    {
                        if (!targetTemplateDict.TryGetValue(targetTemplateIndex, out var controlTemplate))
                            return; // 该靶盘位Index不存在

                        // 多靶多孔
                        if (targetTemplateIndexCount > 1)
                        {
                            var targetVolumeEach = volumeEachList.Count > 1 ? volumeEachList[i] : volumeEachList[0];
                            var targetVolumeTotal = targetVolumeEach.Original * sourceHoleIndexList.Count;
                            controlTemplate.RefreshTemplateHolesStatus(headUsedIndex, targetHoleIndexList[i], tipChannel, targetVolumeTotal);
                        }
                        // 1靶多孔
                        else
                        {
                            for (var j = 0; j < targetHoleIndexList.Count; j++)
                            {
                                var targetHoleIndex = targetHoleIndexList[j];
                                var targetVolumeEach = volumeEachList.Count > 1 ? volumeEachList[j] : volumeEachList[0];
                                var targetVolumeTotal = targetVolumeEach.Original * sourceHoleIndexList.Count;
                                controlTemplate.RefreshTemplateHolesStatus(headUsedIndex, targetHoleIndex, tipChannel, targetVolumeTotal);
                            }
                        }
                    }
                }
            }
        }


        /// <summary>
        /// 重新计算吸液、喷液位置（仅适用于灵活取枪头，如不需要灵活取枪头，可不调用）
        /// </summary>
        private void AllocateAbsorbAndJetInfo()
        {
            for (var i = 0; i < seqList.Count; i++)
            {
                var seq = seqList[i];

                if (seq.IsCmdOnly)
                    continue;

                var headUsedIndex = seq.HeadUsedIndex;

                // 移液头
                var head = ParamsHelper.HeadList[headUsedIndex];

                // 枪头盒耗材
                var tipTemplateConsumableType = seq.TipTemplateConsumableType;
                var tipTemplateConsumableHoleStep = tipTemplateConsumableType.HoleStep;

                // 所需枪头行、列数
                var tipChannelRow = seq.TipChannel.GetLength(0);
                var tipChannelCol = seq.TipChannel.GetLength(1);

                // 源盘耗材
                var sourceTemplateConsumableType = seq.SourceTemplateConsumableType;
                var sourceTemplateConsumableRow = sourceTemplateConsumableType.RowCount;
                var sourceTemplateConsumableCol = sourceTemplateConsumableType.ColCount;
                var sourceTemplateConsumableHoleStep = sourceTemplateConsumableType.HoleStep;
                var sourceHoleIndex = seq.SourceHoleIndexList[0];

                // 靶盘耗材
                var targetTemplateConsumableType = seq.TargetTemplateConsumableType;
                var targetTemplateConsumableRow = targetTemplateConsumableType.RowCount;
                var targetTemplateConsumableCol = targetTemplateConsumableType.ColCount;
                var targetTemplateConsumableHoleStep = targetTemplateConsumableType.HoleStep;

                // 移液头间距与耗材是否不一致
                // var tipHoleStepNotSameAsSourceX = tipChannelCol > 1 && sourceTemplateConsumableHoleStep.X > 0 && tipTemplateConsumableHoleStep.X != sourceTemplateConsumableHoleStep.X;
                // var tipHoleStepNotSameAsSourceY = tipChannelRow > 1 && sourceTemplateConsumableHoleStep.Y > 0 && tipTemplateConsumableHoleStep.Y != sourceTemplateConsumableHoleStep.Y;
                // var tipHoleStepNotSameAsTargetX = tipChannelCol > 1 && targetTemplateConsumableHoleStep.X > 0 && tipTemplateConsumableHoleStep.X != targetTemplateConsumableHoleStep.X;
                // var tipHoleStepNotSameAsTargetY = tipChannelRow > 1 && targetTemplateConsumableHoleStep.Y > 0 && tipTemplateConsumableHoleStep.Y != targetTemplateConsumableHoleStep.Y;
                var tipHoleStepNotSameAsSourceX = head.ChannelCol > 1 && sourceTemplateConsumableHoleStep.X > 0 && tipTemplateConsumableHoleStep.X != sourceTemplateConsumableHoleStep.X;
                var tipHoleStepNotSameAsSourceY = head.ChannelRow > 1 && sourceTemplateConsumableHoleStep.Y > 0 && tipTemplateConsumableHoleStep.Y != sourceTemplateConsumableHoleStep.Y;
                var tipHoleStepNotSameAsTargetX = head.ChannelCol > 1 && targetTemplateConsumableHoleStep.X > 0 && tipTemplateConsumableHoleStep.X != targetTemplateConsumableHoleStep.X;
                var tipHoleStepNotSameAsTargetY = head.ChannelRow > 1 && targetTemplateConsumableHoleStep.Y > 0 && tipTemplateConsumableHoleStep.Y != targetTemplateConsumableHoleStep.Y;

                // y轴移动系数（正方向：头移动；反方向：盘移动）
                var yDirectionFactor = ParamsHelper.HeadList[headUsedIndex].YMoveWithHead ? 1 : -1;

                // 移液头通道数
                var headChannelRow = ParamsHelper.HeadList[headUsedIndex].ChannelRow;
                var headChannelCol = ParamsHelper.HeadList[headUsedIndex].ChannelCol;
                // A1摆放位置
                var a1Pos = ParamsHelper.CommonSettingList[headUsedIndex].A1Pos;
                // 逐列取枪头
                var takeTipEachCol = ParamsHelper.CommonSettingList[headUsedIndex].TakeTipEachCol;
                // 取枪头方向：从左往右
                var takeTipLeft2Right = ParamsHelper.CommonSettingList[headUsedIndex].TakeTipLeft2Right || headChannelRow * headChannelCol == 1;

                // 是否灵活取枪头
                bool isTakeTipFlexible = TipHelper.IsTakeTipFlexible(headUsedIndex, seq.TipChannel, tipTemplateConsumableType);

                // 是否灵活取枪头
                if (isTakeTipFlexible)
                {
                    // A1左上
                    if (a1Pos == EA1Pos.LeftTop)
                    {
                        // 逐列取
                        if (takeTipEachCol)
                        {
                            /**
                             * 源孔
                             */
                            // 所在列Index
                            var sourceColIndex = sourceHoleIndex.OriIndex / sourceTemplateConsumableRow * sourceTemplateConsumableRow;

                            /**
                            * 移液头偏移逻辑：
                            * ①移液头多列，整列取：X轴偏移
                            * ②移液头多列，灵活取：X、Y轴偏移
                            * ③移液头单列，Y轴偏移
                            */
                            // 偏移 = 取枪头数目Col - 移液头通道Col
                            var sourceXHoleOffset = takeTipLeft2Right ? tipChannelCol - headChannelCol : 0;
                            // 偏移 = holeIndex + 取枪头数目Row -（所在列首孔Index + 移液头通道Row）
                            var sourceYHoleOffset = (sourceHoleIndex.OriIndex + tipChannelRow - (sourceColIndex + headChannelRow)) * yDirectionFactor;
                            sourceHoleIndex.XHoleOffset = sourceXHoleOffset;
                            sourceHoleIndex.YHoleOffset = sourceYHoleOffset;

                            // 移液头通道间距与耗材间距是否不一致
                            if (!head.IsVariable)
                            {
                                // 单通道
                                if (tipChannelRow == 1)
                                {
                                    sourceHoleIndex.StepNotSameX = tipHoleStepNotSameAsSourceX;
                                    sourceHoleIndex.StepNotSameY = tipHoleStepNotSameAsSourceY;
                                    // 特殊处理：2025-10-18， StepNotSameX和StepNotSameY均为false，则认为间距一致，OriIndex置为列index即可
                                    if (!sourceHoleIndex.StepNotSameX && !sourceHoleIndex.StepNotSameY)
                                        sourceHoleIndex.OriIndex = sourceColIndex;

                                    // 特殊处理：例如溶液槽x孔距比移液头x孔距大，默认走到溶液槽A1，不走偏移
                                    // （暂舍弃 2024-03-18）
                                    // if (stepNotSameAsSourceX && sourceTemplateConsumableHoleStep.X > tipTemplateConsumableHoleStep.X && sourceTemplateConsumableRow == 1)
                                    // {
                                    //     seq.SourceHoleIndex.XHoleOffset = 0;
                                    //     seq.SourceHoleIndex.YHoleOffset = 0;
                                    // }
                                    // 特殊处理： 移液头通道间距与耗材间距不一致下，耗材row数大于移液头row数，YHoleOffset默认为负数
                                    if (sourceHoleIndex.StepNotSameY && sourceTemplateConsumableRow > headChannelRow)
                                        sourceHoleIndex.YHoleOffset = -1;
                                }
                                // 多通道
                                else
                                {
                                    // 特殊处理：例如溶液槽x孔距比移液头x孔距大，默认走到溶液槽A1，不走偏移
                                    if (tipChannelRow > sourceTemplateConsumableRow)
                                    {
                                        var stepNotSameAsSourceX = sourceHoleIndex.XHoleOffset != 0 && tipHoleStepNotSameAsSourceX;
                                        sourceHoleIndex.StepNotSameX = stepNotSameAsSourceX;
                                        // （暂舍弃 2024-03-18）
                                        // if (stepNotSameAsSourceX && sourceTemplateConsumableHoleStep.X > tipTemplateConsumableHoleStep.X && sourceTemplateConsumableRow == 1)
                                        // {
                                        //     seq.SourceHoleIndex.XHoleOffset = 0;
                                        //     seq.SourceHoleIndex.YHoleOffset = 0;
                                        // }
                                    }
                                    // 如果移液头与耗材行数一致，OriIndex变为所在列Index
                                    else if (headChannelRow == sourceTemplateConsumableRow)
                                    {
                                        sourceHoleIndex.OriIndex = sourceColIndex;
                                        sourceHoleIndex.StepNotSameX = tipHoleStepNotSameAsSourceX;
                                    }

                                    // 特殊处理：如果行距 == 0 或者 行距 >= 移液头行距和，则认为是溶液槽，y不走偏移（2024-05-28）
                                    if (sourceTemplateConsumableHoleStep.Y == 0 || sourceTemplateConsumableHoleStep.Y >= headChannelRow * tipTemplateConsumableHoleStep.Y)
                                    {
                                        sourceHoleIndex.YHoleOffset = 0;
                                        sourceHoleIndex.StepNotSameY = false;
                                    }
                                }
                            }

                            /**
                             * 靶孔
                             */
                            foreach (var targetHoleIndex in seq.TargetHoleIndexList)
                            {
                                // 所在列Index
                                var targetColIndex = targetHoleIndex.OriIndex / targetTemplateConsumableRow * targetTemplateConsumableRow;

                                /**
                                * 移液头偏移逻辑：
                                * ①移液头多列，整列取：X轴偏移
                                * ②移液头多列，灵活取：X、Y轴偏移
                                * ③移液头单列，Y轴偏移
                                */
                                // 偏移 = 取枪头数目Col - 移液头通道Col
                                var targetXHoleOffset = takeTipLeft2Right ? tipChannelCol - headChannelCol : 0;
                                // 偏移 = holeIndex + 取枪头数目Row -（所在列Index + 移液头通道Row）
                                var targetYHoleOffset = (targetHoleIndex.OriIndex + tipChannelRow - (targetColIndex + headChannelRow)) * yDirectionFactor;
                                targetHoleIndex.XHoleOffset = targetXHoleOffset;
                                targetHoleIndex.YHoleOffset = targetYHoleOffset;

                                // 移液头通道间距与耗材间距是否不一致
                                if (!head.IsVariable)
                                {
                                    // 单通道
                                    if (tipChannelRow == 1)
                                    {
                                        // if (targetTemplateConsumableHoleStep.Y != head.ChannelStep)
                                        targetHoleIndex.StepNotSameX = tipHoleStepNotSameAsTargetX;
                                        targetHoleIndex.StepNotSameY = tipHoleStepNotSameAsTargetY;
                                        // 特殊处理：2025-10-18， StepNotSameX和StepNotSameY均为false，则认为间距一致，OriIndex置为列index即可
                                        if (!targetHoleIndex.StepNotSameX && !targetHoleIndex.StepNotSameY)
                                            targetHoleIndex.OriIndex = targetColIndex;

                                        // 特殊处理：例如溶液槽x孔距比移液头x孔距大，默认走到溶液槽A1，不走偏移
                                        // （暂舍弃 2024-03-18）
                                        // var stepNotSameAsTargetX = targetHoleIndex.XHoleOffset != 0 && tipHoleStepNotSameAsTargetX;
                                        // if (stepNotSameAsTargetX && targetTemplateConsumableHoleStep.X > tipTemplateConsumableHoleStep.X && targetTemplateConsumableRow == 1)
                                        // {
                                        //     targetHoleIndex.XHoleOffset = 0;
                                        //     targetHoleIndex.YHoleOffset = 0;
                                        // }
                                        // 特殊处理： 移液头通道间距与耗材间距不一致下，耗材row数大于移液头row数，YHoleOffset默认为负数
                                        if (targetHoleIndex.StepNotSameY && targetTemplateConsumableRow > headChannelRow)
                                            targetHoleIndex.YHoleOffset = -1;
                                    }
                                    // 多通道
                                    else
                                    {
                                        // 特殊处理：例如溶液槽x孔距比移液头x孔距大，默认走到溶液槽A1，不走偏移
                                        if (tipChannelRow > targetTemplateConsumableRow)
                                        {
                                            var stepNotSameAsTargetX = targetHoleIndex.XHoleOffset != 0 && tipHoleStepNotSameAsTargetX;
                                            sourceHoleIndex.StepNotSameX = stepNotSameAsTargetX;
                                            // （暂舍弃 2024-03-18）
                                            // if (stepNotSameAsTargetX && targetTemplateConsumableHoleStep.X > tipTemplateConsumableHoleStep.X && targetTemplateConsumableRow == 1)
                                            // {
                                            //     targetHoleIndex.XHoleOffset = 0;
                                            //     targetHoleIndex.YHoleOffset = 0;
                                            // }
                                        }
                                        // 如果移液头与耗材行数一致，OriIndex变为所在列Index
                                        else if (headChannelRow == targetTemplateConsumableRow)
                                        {
                                            targetHoleIndex.OriIndex = targetColIndex;
                                            targetHoleIndex.StepNotSameX = tipHoleStepNotSameAsTargetX;
                                        }

                                        // 特殊处理：如果行距 == 0 或者 行距 >= 移液头行距和，则认为是溶液槽，y不走偏移（2024-05-28）
                                        if (targetTemplateConsumableHoleStep.Y == 0 || targetTemplateConsumableHoleStep.Y >= headChannelRow * tipTemplateConsumableHoleStep.Y)
                                        {
                                            targetHoleIndex.YHoleOffset = 0;
                                            targetHoleIndex.StepNotSameY = false;
                                        }
                                    }
                                }
                            }
                        }
                        // 逐行取
                        else
                        {
                            /**
                              * 源孔
                              */
                            // 所在行Index
                            var sourceRowIndex = sourceHoleIndex.OriIndex % sourceTemplateConsumableRow;
                            // 所在列Index
                            var sourceColIndex = sourceHoleIndex.OriIndex / sourceTemplateConsumableRow;

                            /**
                            * 移液头偏移逻辑：
                            * ①移液头多行，整行取：Y轴偏移
                            * ②移液头多行，灵活取：X、Y轴偏移
                            * ③移液头单行，X轴偏移
                            */
                            // 偏移 = 取枪头数目Col -（移液头通道Col - 所在列Index）
                            var sourceXHoleOffset = takeTipLeft2Right ? tipChannelCol - (headChannelCol - sourceColIndex) : 0;
                            // 偏移 = 取枪头数目Row - 移液头通道Row
                            var sourceYHoleOffset = (tipChannelRow - headChannelRow) * yDirectionFactor;
                            sourceHoleIndex.XHoleOffset = sourceXHoleOffset;
                            sourceHoleIndex.YHoleOffset = sourceYHoleOffset;

                            // 移液头通道间距与耗材间距是否不一致
                            if (!head.IsVariable)
                            {
                                // 单通道
                                if (tipChannelCol == 1)
                                {
                                    // if (sourceTemplateConsumableHoleStep.X != head.ChannelStep)
                                    sourceHoleIndex.StepNotSameX = tipHoleStepNotSameAsSourceX;
                                    sourceHoleIndex.StepNotSameY = tipHoleStepNotSameAsSourceY;
                                    // 特殊处理：2025-10-18， StepNotSameX和StepNotSameY均为false，则认为间距一致，OriIndex置为行index即可
                                    if (!sourceHoleIndex.StepNotSameX && !sourceHoleIndex.StepNotSameY) 
                                        sourceHoleIndex.OriIndex = sourceRowIndex;
                                    // 特殊处理：例如溶液槽y孔距比移液头y孔距大，默认走到溶液槽A1，不走偏移
                                    // （暂舍弃 2024-03-18）
                                    // var stepNotSameAsSourceY = seq.SourceHoleIndex.YHoleOffset != 0 && tipHoleStepNotSameAsSourceY;
                                    // if (stepNotSameAsSourceY && sourceTemplateConsumableHoleStep.Y > tipTemplateConsumableHoleStep.Y && sourceTemplateConsumableCol == 1)
                                    // {
                                    //     seq.SourceHoleIndex.XHoleOffset = 0;
                                    //     seq.SourceHoleIndex.YHoleOffset = 0;
                                    // }
                                    // 特殊处理： 移液头通道间距与耗材间距不一致下，耗材col数大于移液头col数，XHoleOffset默认为负数
                                    if (sourceHoleIndex.StepNotSameX && sourceTemplateConsumableCol > headChannelCol)
                                        sourceHoleIndex.XHoleOffset = -1;
                                }
                                // 多通道
                                else
                                {
                                    // 特殊处理：例如溶液槽y孔距比移液头y孔距大，默认走到溶液槽A1，不走偏移
                                    if (tipChannelCol > sourceTemplateConsumableCol)
                                    {
                                        var stepNotSameAsSourceY = sourceHoleIndex.YHoleOffset != 0 && tipHoleStepNotSameAsSourceY;
                                        sourceHoleIndex.StepNotSameY = stepNotSameAsSourceY;
                                        // （暂舍弃 2024-03-18）
                                        // if (stepNotSameAsSourceY && sourceTemplateConsumableHoleStep.Y > tipTemplateConsumableHoleStep.Y && sourceTemplateConsumableCol == 1)
                                        // {
                                        //     seq.SourceHoleIndex.XHoleOffset = 0;
                                        //     seq.SourceHoleIndex.YHoleOffset = 0;
                                        // }
                                    }
                                    // 如果移液头与耗材列数一致，OriIndex变为所在行Index
                                    else if (headChannelCol == sourceTemplateConsumableCol)
                                    {
                                        sourceHoleIndex.OriIndex = sourceRowIndex;
                                        sourceHoleIndex.StepNotSameY = tipHoleStepNotSameAsSourceY;
                                    }

                                    // 特殊处理：如果列距 == 0 或者 列距 >= 移液头列距和，则认为是溶液槽，x不走偏移（2024-05-28）
                                    if (sourceTemplateConsumableHoleStep.X == 0 || sourceTemplateConsumableHoleStep.X >= headChannelCol * tipTemplateConsumableHoleStep.X)
                                    {
                                        sourceHoleIndex.XHoleOffset = 0;
                                        sourceHoleIndex.StepNotSameX = false;
                                    }
                                }
                            }

                            /**
                             * 靶孔
                             */
                            foreach (var targetHoleIndex in seq.TargetHoleIndexList)
                            {
                                // 所在行Index
                                var targetRowIndex = targetHoleIndex.OriIndex % targetTemplateConsumableRow;
                                // 所在列Index
                                var targetColIndex = targetHoleIndex.OriIndex / targetTemplateConsumableRow;

                                /**
                                * 移液头偏移逻辑：
                                * ①移液头多行，但整行取：Y轴偏移
                                * ②移液头多行，但灵活列取：X、Y轴偏移
                                * ③移液头单行，X轴偏移
                                */
                                // 偏移 = 取枪头数目Col -（移液头通道Col - 所在列Index）
                                var targetXHoleOffset = takeTipLeft2Right ? tipChannelCol - (headChannelCol - targetColIndex) : 0;
                                // 偏移 = 取枪头数目Row - 移液头通道Row
                                var targetYHoleOffset = (tipChannelRow - headChannelRow) * yDirectionFactor;
                                targetHoleIndex.XHoleOffset = targetXHoleOffset;
                                targetHoleIndex.YHoleOffset = targetYHoleOffset;

                                // 移液头通道间距与耗材间距是否不一致
                                if (!head.IsVariable)
                                {
                                    // 单通道
                                    if (tipChannelCol == 1)
                                    {
                                        // if (targetTemplateConsumableHoleStep.X != head.ChannelStep)
                                        targetHoleIndex.StepNotSameX = tipHoleStepNotSameAsTargetX;
                                        targetHoleIndex.StepNotSameY = tipHoleStepNotSameAsTargetY;
                                        // 特殊处理：2025-10-18， StepNotSameX和StepNotSameY均为false，则认为间距一致，OriIndex置为行index即可
                                        if (!targetHoleIndex.StepNotSameX && !targetHoleIndex.StepNotSameY)
                                            targetHoleIndex.OriIndex = targetRowIndex;
                                        // 特殊处理：例如溶液槽y孔距比移液头y孔距大，默认走到溶液槽A1，不走偏移
                                        // （暂舍弃 2024-03-18）
                                        // var stepNotSameAsTargetY = targetHoleIndex.YHoleOffset != 0 && tipHoleStepNotSameAsTargetY;
                                        // if (stepNotSameAsTargetY && targetTemplateConsumableHoleStep.Y > tipTemplateConsumableHoleStep.Y && targetTemplateConsumableCol == 1)
                                        // {
                                        //     targetHoleIndex.XHoleOffset = 0;
                                        //     targetHoleIndex.YHoleOffset = 0;
                                        // }
                                        // 特殊处理： 移液头通道间距与耗材间距不一致下，耗材col数大于移液头col数，XHoleOffset默认为负数
                                        if (targetHoleIndex.StepNotSameX && targetTemplateConsumableCol > headChannelCol)
                                            targetHoleIndex.XHoleOffset = -1;
                                    }
                                    // 多通道
                                    else
                                    {
                                        // 特殊处理：例如溶液槽y孔距比移液头y孔距大，默认走到溶液槽A1，不走偏移
                                        if (tipChannelCol > targetTemplateConsumableCol)
                                        {
                                            var stepNotSameAsTargetY = targetHoleIndex.YHoleOffset != 0 && tipHoleStepNotSameAsTargetY;
                                            targetHoleIndex.StepNotSameY = stepNotSameAsTargetY;
                                            // （暂舍弃 2024-03-18）
                                            // if (stepNotSameAsTargetY && targetTemplateConsumableHoleStep.Y > tipTemplateConsumableHoleStep.Y && targetTemplateConsumableCol == 1)
                                            // {
                                            //     targetHoleIndex.XHoleOffset = 0;
                                            //     targetHoleIndex.YHoleOffset = 0;
                                            // }
                                        }
                                        // 如果移液头与耗材列数一致，OriIndex变为所在行Index
                                        else if (headChannelCol == targetTemplateConsumableCol)
                                        {
                                            targetHoleIndex.OriIndex = targetRowIndex;
                                            targetHoleIndex.StepNotSameY = tipHoleStepNotSameAsTargetY;
                                        }

                                        // 特殊处理：如果列距 == 0 或者 列距 >= 移液头列距和，则认为是溶液槽，x不走偏移（2024-05-28）
                                        if (targetTemplateConsumableHoleStep.X == 0 || targetTemplateConsumableHoleStep.X >= headChannelCol * tipTemplateConsumableHoleStep.X)
                                        {
                                            targetHoleIndex.XHoleOffset = 0;
                                            targetHoleIndex.StepNotSameX = false;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    // A1左下
                    else
                    {
                        // 逐列取
                        if (takeTipEachCol)
                        {
                            /**
                              * 源孔
                              */
                            // 所在行Index
                            var sourceRowIndex = sourceHoleIndex.OriIndex / sourceTemplateConsumableCol;
                            // 所在列Index
                            var sourceColIndex = sourceHoleIndex.OriIndex % sourceTemplateConsumableCol;

                            /**
                            * 移液头偏移逻辑：
                            * ①移液头多列，整列取：X轴偏移
                            * ②移液头多列，灵活取：X、Y轴偏移
                            * ③移液头单列，Y轴偏移
                            */
                            // 偏移 = 取枪头数目Col - 移液头通道Col
                            var sourceXHoleOffset = takeTipLeft2Right ? tipChannelCol - headChannelCol : 0;
                            // 偏移 = 取枪头数目Row -（移液头通道Row - 所在行Index）
                            var sourceYHoleOffset = (headChannelRow - sourceRowIndex - tipChannelRow) * yDirectionFactor;
                            sourceHoleIndex.XHoleOffset = sourceXHoleOffset;
                            sourceHoleIndex.YHoleOffset = sourceYHoleOffset;

                            // 移液头通道间距与耗材间距是否不一致
                            if (!head.IsVariable)
                            {
                                // 单通道
                                if (tipChannelRow == 1)
                                {
                                    // if (sourceTemplateConsumableHoleStep.Y != head.ChannelStep)
                                    sourceHoleIndex.StepNotSameX = tipHoleStepNotSameAsSourceX;
                                    sourceHoleIndex.StepNotSameY = tipHoleStepNotSameAsSourceY;
                                    // 特殊处理：2025-10-18， StepNotSameX和StepNotSameY均为false，则认为间距一致，OriIndex置为列index即可
                                    if (!sourceHoleIndex.StepNotSameX && !sourceHoleIndex.StepNotSameY)
                                        sourceHoleIndex.OriIndex = sourceColIndex;
                                    // 特殊处理：例如溶液槽x孔距比移液头x孔距大，默认走到溶液槽A1，不走偏移
                                    // （暂舍弃 2024-03-18）
                                    // var stepNotSameAsSourceX = seq.SourceHoleIndex.XHoleOffset != 0 && tipHoleStepNotSameAsSourceX;
                                    // if (stepNotSameAsSourceX && sourceTemplateConsumableHoleStep.X > tipTemplateConsumableHoleStep.X && sourceTemplateConsumableRow == 1)
                                    // {
                                    //     seq.SourceHoleIndex.XHoleOffset = 0;
                                    //     seq.SourceHoleIndex.YHoleOffset = 0;
                                    // }
                                    // 特殊处理： 移液头通道间距与耗材间距不一致下，耗材row数大于移液头row数，YHoleOffset默认为负数
                                    if (sourceHoleIndex.StepNotSameY && sourceTemplateConsumableRow > headChannelRow)
                                        sourceHoleIndex.YHoleOffset = -1;
                                }
                                // 多通道
                                else
                                {
                                    // 特殊处理：例如溶液槽x孔距比移液头x孔距大，默认走到溶液槽A1，不走偏移
                                    if (tipChannelRow > sourceTemplateConsumableRow)
                                    {
                                        var stepNotSameAsSourceX = sourceHoleIndex.XHoleOffset != 0 && tipHoleStepNotSameAsSourceX;
                                        sourceHoleIndex.StepNotSameX = stepNotSameAsSourceX;
                                        // （暂舍弃 2024-03-18）
                                        // if (stepNotSameAsSourceX && sourceTemplateConsumableHoleStep.X > tipTemplateConsumableHoleStep.X && sourceTemplateConsumableRow == 1)
                                        // {
                                        //     seq.SourceHoleIndex.XHoleOffset = 0;
                                        //     seq.SourceHoleIndex.YHoleOffset = 0;
                                        // }
                                    }
                                    // 如果移液头与耗材行数一致，OriIndex变为所在列Index
                                    else if (headChannelRow == sourceTemplateConsumableRow)
                                    {
                                        sourceHoleIndex.OriIndex = sourceColIndex;
                                        sourceHoleIndex.StepNotSameX = tipHoleStepNotSameAsSourceX;
                                    }

                                    // 特殊处理：如果行距 == 0 或者 行距 >= 移液头行距和，则认为是溶液槽，y不走偏移（2024-05-28）
                                    if (sourceTemplateConsumableHoleStep.Y == 0 || sourceTemplateConsumableHoleStep.Y >= headChannelRow * tipTemplateConsumableHoleStep.Y)
                                    {
                                        sourceHoleIndex.YHoleOffset = 0;
                                        sourceHoleIndex.StepNotSameY = false;
                                    }
                                }
                            }

                            /**
                             * 靶孔
                             */
                            foreach (var targetHoleIndex in seq.TargetHoleIndexList)
                            {
                                // 所在行Index
                                var targetRowIndex = targetHoleIndex.OriIndex / targetTemplateConsumableCol;
                                // 所在列Index
                                var targetColIndex = targetHoleIndex.OriIndex % targetTemplateConsumableCol;

                                /**
                                * 移液头偏移逻辑：
                                * ①移液头多列，整列取：X轴偏移
                                * ②移液头多列，灵活取：X、Y轴偏移
                                * ③移液头单列，Y轴偏移
                                */
                                // 偏移 = 取枪头数目Col - 移液头通道Col
                                var targetXHoleOffset = takeTipLeft2Right ? tipChannelCol - headChannelCol : 0;
                                // 偏移 = 取枪头数目Row -（移液头通道Row - 所在行Index）
                                var targetYHoleOffset = (headChannelRow - targetRowIndex - tipChannelRow) * yDirectionFactor;
                                targetHoleIndex.XHoleOffset = targetXHoleOffset;
                                targetHoleIndex.YHoleOffset = targetYHoleOffset;

                                // 移液头通道间距与耗材间距是否不一致
                                if (!head.IsVariable)
                                {
                                    // 单通道
                                    if (tipChannelRow == 1)
                                    {
                                        // if (targetTemplateConsumableHoleStep.Y != head.ChannelStep)
                                        targetHoleIndex.StepNotSameX = tipHoleStepNotSameAsTargetX;
                                        targetHoleIndex.StepNotSameY = tipHoleStepNotSameAsTargetY;
                                        // 特殊处理：2025-10-18， StepNotSameX和StepNotSameY均为false，则认为间距一致，OriIndex置为列index即可
                                        if (!targetHoleIndex.StepNotSameX && !targetHoleIndex.StepNotSameY)
                                            targetHoleIndex.OriIndex = targetColIndex;
                                        // 特殊处理：例如溶液槽x孔距比移液头x孔距大，默认走到溶液槽A1，不走偏移
                                        // （暂舍弃 2024-03-18）
                                        // var stepNotSameAsTargetX = targetHoleIndex.XHoleOffset != 0 && tipHoleStepNotSameAsTargetX;
                                        // if (stepNotSameAsTargetX && targetTemplateConsumableHoleStep.X > tipTemplateConsumableHoleStep.X && targetTemplateConsumableRow == 1)
                                        // {
                                        //     targetHoleIndex.XHoleOffset = 0;
                                        //     targetHoleIndex.YHoleOffset = 0;
                                        // }
                                        // 特殊处理： 移液头通道间距与耗材间距不一致下，耗材row数大于移液头row数，YHoleOffset默认为负数
                                        if (targetHoleIndex.StepNotSameY && targetTemplateConsumableRow > headChannelRow)
                                            targetHoleIndex.YHoleOffset = -1;
                                    }
                                    // 多通道
                                    else
                                    {
                                        // 特殊处理：例如溶液槽x孔距比移液头x孔距大，默认走到溶液槽A1，不走偏移
                                        if (tipChannelRow > targetTemplateConsumableRow)
                                        {
                                            var stepNotSameAsTargetX = targetHoleIndex.XHoleOffset != 0 && tipHoleStepNotSameAsTargetX;
                                            sourceHoleIndex.StepNotSameX = stepNotSameAsTargetX;
                                            // （暂舍弃 2024-03-18）
                                            // if (stepNotSameAsTargetX && targetTemplateConsumableHoleStep.X > tipTemplateConsumableHoleStep.X && targetTemplateConsumableRow == 1)
                                            // {
                                            //     targetHoleIndex.XHoleOffset = 0;
                                            //     targetHoleIndex.YHoleOffset = 0;
                                            // }
                                        }
                                        // 如果移液头与耗材行数一致，OriIndex变为所在列Index
                                        else if (headChannelRow == targetTemplateConsumableRow)
                                        {
                                            targetHoleIndex.OriIndex = targetColIndex;
                                            targetHoleIndex.StepNotSameX = tipHoleStepNotSameAsTargetX;
                                        }

                                        // 特殊处理：如果行距 == 0 或者 行距 >= 移液头行距和，则认为是溶液槽，y不走偏移（2024-05-28）
                                        if (targetTemplateConsumableHoleStep.Y == 0 || targetTemplateConsumableHoleStep.Y >= headChannelRow * tipTemplateConsumableHoleStep.Y)
                                        {
                                            targetHoleIndex.YHoleOffset = 0;
                                            targetHoleIndex.StepNotSameY = false;
                                        }
                                    }
                                }
                            }
                        }
                        // 逐行取
                        else
                        {
                            /**
                             * 源孔
                             */
                            // 所在行首孔Index
                            var sourceRowFirstIndex = sourceHoleIndex.OriIndex / sourceTemplateConsumableCol * sourceTemplateConsumableCol;
                            // 所在列Index
                            var sourceColIndex = sourceHoleIndex.OriIndex % sourceTemplateConsumableCol;

                            /**
                            * 移液头偏移逻辑：
                            * ①移液头多行，整行取：Y轴偏移
                            * ②移液头多行，灵活取：X、Y轴偏移
                            * ③移液头单行，X轴偏移
                            */
                            // 偏移 = 所在列index + 取枪头数目Col - 移液头通道Col
                            var sourceXHoleOffset = takeTipLeft2Right ? sourceColIndex + tipChannelCol - headChannelCol : 0;
                            // 偏移 = 移液头通道Row - 取枪头数目Row
                            var sourceYHoleOffset = (headChannelRow - tipChannelRow) * yDirectionFactor;
                            sourceHoleIndex.XHoleOffset = sourceXHoleOffset;
                            sourceHoleIndex.YHoleOffset = sourceYHoleOffset;

                            // 移液头通道间距与耗材间距是否不一致
                            if (!head.IsVariable)
                            {
                                // 单通道
                                if (tipChannelCol == 1)
                                {
                                    // if (sourceTemplateConsumableHoleStep.X != head.ChannelStep)
                                    sourceHoleIndex.StepNotSameX = tipHoleStepNotSameAsSourceX;
                                    sourceHoleIndex.StepNotSameY = tipHoleStepNotSameAsSourceY;
                                    // if (!sourceHoleIndex.StepNotSameX || !sourceHoleIndex.StepNotSameY)
                                    //     sourceHoleIndex.OriIndex = sourceRowFirstIndex;

                                    // 特殊处理：例如溶液槽y孔距比移液头y孔距大，默认走到溶液槽A1，不走偏移
                                    // （暂舍弃 2024-03-18）
                                    // var stepNotSameAsSourceY = seq.SourceHoleIndex.YHoleOffset != 0 && tipHoleStepNotSameAsSourceY;
                                    // if (stepNotSameAsSourceY && sourceTemplateConsumableHoleStep.Y > tipTemplateConsumableHoleStep.Y && sourceTemplateConsumableCol == 1)
                                    // {
                                    //     seq.SourceHoleIndex.XHoleOffset = 0;
                                    //     seq.SourceHoleIndex.YHoleOffset = 0;
                                    // }
                                    // 特殊处理： 移液头通道间距与耗材间距不一致下，耗材col数大于移液头col数，XHoleOffset默认为负数
                                    if (sourceHoleIndex.StepNotSameX && sourceTemplateConsumableCol > headChannelCol)
                                        sourceHoleIndex.XHoleOffset = -1;
                                }
                                // 多通道
                                else
                                {
                                    // 特殊处理：例如溶液槽y孔距比移液头y孔距大，默认走到溶液槽A1，不走偏移
                                    if (tipChannelCol > sourceTemplateConsumableCol)
                                    {
                                        var stepNotSameAsSourceY = sourceHoleIndex.YHoleOffset != 0 && tipHoleStepNotSameAsSourceY;
                                        sourceHoleIndex.StepNotSameY = stepNotSameAsSourceY;
                                        // （暂舍弃 2024-03-18）
                                        // if (stepNotSameAsSourceY && sourceTemplateConsumableHoleStep.Y > tipTemplateConsumableHoleStep.Y && sourceTemplateConsumableCol == 1)
                                        // {
                                        //     seq.SourceHoleIndex.XHoleOffset = 0;
                                        //     seq.SourceHoleIndex.YHoleOffset = 0;
                                        // }
                                    }
                                    // 如果移液头与耗材列数一致，OriIndex变为所在行Index
                                    else if (headChannelCol == sourceTemplateConsumableCol)
                                    {
                                        sourceHoleIndex.OriIndex = sourceRowFirstIndex;
                                        sourceHoleIndex.StepNotSameY = tipHoleStepNotSameAsSourceY;
                                    }

                                    // 特殊处理：如果列距 == 0 或者 列距 >= 移液头列距和，则认为是溶液槽，x不走偏移（2024-05-28）
                                    if (sourceTemplateConsumableHoleStep.X == 0 || sourceTemplateConsumableHoleStep.X >= headChannelCol * tipTemplateConsumableHoleStep.X)
                                    {
                                        sourceHoleIndex.XHoleOffset = 0;
                                        sourceHoleIndex.StepNotSameX = false;
                                    }
                                }
                            }

                            /**
                             * 靶孔
                             */
                            foreach (var targetHoleIndex in seq.TargetHoleIndexList)
                            {
                                // 所在行首孔Index
                                var targetRowFirstIndex = targetHoleIndex.OriIndex / targetTemplateConsumableCol * targetTemplateConsumableCol;
                                // 所在列ndex
                                var targetColIndex = targetHoleIndex.OriIndex % targetTemplateConsumableCol;

                                /**
                                * 移液头偏移逻辑：
                                * ①移液头多行，整行取：Y轴偏移
                                * ②移液头多行，灵活取：X、Y轴偏移
                                * ③移液头单行，X轴偏移
                                */
                                // 偏移 = 所在列index + 取枪头数目Col - 移液头通道Col
                                var targetXHoleOffset = takeTipLeft2Right ? targetColIndex + tipChannelCol - headChannelCol : 0;
                                // 偏移 = 移液头通道Row - 取枪头数目Row
                                var targetYHoleOffset = (headChannelRow - tipChannelRow) * yDirectionFactor;
                                targetHoleIndex.XHoleOffset = targetXHoleOffset;
                                targetHoleIndex.YHoleOffset = targetYHoleOffset;

                                // 移液头通道间距与耗材间距是否不一致
                                if (!head.IsVariable)
                                {
                                    // 单通道
                                    if (tipChannelCol == 1)
                                    {
                                        // if (targetTemplateConsumableHoleStep.X != head.ChannelStep)
                                        targetHoleIndex.StepNotSameX = tipHoleStepNotSameAsTargetX;
                                        targetHoleIndex.StepNotSameY = tipHoleStepNotSameAsTargetY;
                                        // if (!targetHoleIndex.StepNotSameX || !targetHoleIndex.StepNotSameY)
                                        //     targetHoleIndex.OriIndex = targetRowFirstIndex;

                                        // 特殊处理：例如溶液槽y孔距比移液头y孔距大，默认走到溶液槽A1，不走偏移
                                        // （暂舍弃 2024-03-18）
                                        // var stepNotSameAsTargetY = targetHoleIndex.YHoleOffset != 0 && tipHoleStepNotSameAsTargetY;
                                        // if (stepNotSameAsTargetY && targetTemplateConsumableHoleStep.Y > tipTemplateConsumableHoleStep.Y && targetTemplateConsumableCol == 1)
                                        // {
                                        //     targetHoleIndex.XHoleOffset = 0;
                                        //     targetHoleIndex.YHoleOffset = 0;
                                        // }
                                        // 特殊处理： 移液头通道间距与耗材间距不一致下，耗材col数大于移液头col数，XHoleOffset默认为负数
                                        if (targetHoleIndex.StepNotSameX && targetTemplateConsumableCol > headChannelCol)
                                            targetHoleIndex.XHoleOffset = -1;
                                    }
                                    // 多通道
                                    else
                                    {
                                        // 特殊处理：例如溶液槽y孔距比移液头y孔距大，默认走到溶液槽A1，不走偏移
                                        if (tipChannelCol > targetTemplateConsumableCol)
                                        {
                                            var stepNotSameAsTargetY = targetHoleIndex.YHoleOffset != 0 && tipHoleStepNotSameAsTargetY;
                                            targetHoleIndex.StepNotSameY = stepNotSameAsTargetY;
                                            // （暂舍弃 2024-03-18）
                                            // if (stepNotSameAsTargetY && targetTemplateConsumableHoleStep.Y > tipTemplateConsumableHoleStep.Y && targetTemplateConsumableCol == 1)
                                            // {
                                            //     targetHoleIndex.XHoleOffset = 0;
                                            //     targetHoleIndex.YHoleOffset = 0;
                                            // }
                                        }
                                        // 如果移液头与耗材列数一致，OriIndex变为所在行Index
                                        else if (headChannelCol == targetTemplateConsumableCol)
                                        {
                                            targetHoleIndex.OriIndex = targetRowFirstIndex;
                                            targetHoleIndex.StepNotSameY = tipHoleStepNotSameAsTargetY;
                                        }

                                        // 特殊处理：如果列距 == 0 或者 列距 >= 移液头列距和，则认为是溶液槽，x不走偏移（2024-05-28）
                                        if (targetTemplateConsumableHoleStep.X == 0 || targetTemplateConsumableHoleStep.X >= headChannelCol * tipTemplateConsumableHoleStep.X)
                                        {
                                            targetHoleIndex.XHoleOffset = 0;
                                            targetHoleIndex.StepNotSameX = false;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 分配退枪头信息
        /// </summary>
        private void AllocateReleaseTipInfo()
        {
            /**
             * 是否退枪头
             */
            for (var i = 0; i < seqList.Count; i++)
            {
                // 当前Seq
                var seqCurrent = seqList[i];

                if (seqCurrent.IsCmdOnly || seqCurrent.IsPumpLiquid)
                    continue;

                // 下一个Seq
                var seqNext = i == seqList.Count - 1 ? null : seqList.ElementAt(i + 1);
                // 特殊指令
                var cmd = seqCurrent.Cmd.ToLower();
                // 最后一个移液信息完成后是否退枪头
                var lastSeqReleaseTip = !(seqNext == null && cmd.Contains(AutoLiquid_Library.Utils.ConstantsUtils.NoReleaseTipCmd.ToLower()));
                // 是否强制退枪头
                var seqReleaseTip = cmd.Contains(AutoLiquid_Library.Utils.ConstantsUtils.ReleaseTipCmd.ToLower()) && !cmd.Contains(AutoLiquid_Library.Utils.ConstantsUtils.NoReleaseTipCmd.ToLower());

                //var isReleaseTip = seqReleaseTip || (seqNext == null && lastSeqReleaseTip) || (seqNext != null && seqNext.IsTakeTip && seqCurrent.IsTakeTip);
                var isReleaseTip = seqReleaseTip || (seqNext == null && lastSeqReleaseTip) || (seqNext != null && seqNext.IsTakeTip);
                seqList[i].IsReleaseTip = isReleaseTip;
            }

            /**
             * 移液头分配退枪头位置
             */
            for (var headIndex = 0; headIndex < ParamsHelper.HeadList.Count; headIndex++)
            {
                // 需要退枪头的seq列表
                var seqListUsedHead = seqList.Where(p => p.HeadUsedIndex == headIndex && p.IsReleaseTip).ToList();
                // 退枪头位置可用Index列表
                var releaseTipPosAvailableIndexList = new List<int>();
                for (var i = 0; i < ParamsHelper.CommonSettingList[headIndex].ReleaseTipPosAvailableList.Count; i++)
                {
                    if (ParamsHelper.CommonSettingList[headIndex].ReleaseTipPosAvailableList[i])
                        releaseTipPosAvailableIndexList.Add(i);
                }
                // 退枪头位置可用数目
                var releaseTipPosAvailableCount = releaseTipPosAvailableIndexList.Count;

                if (releaseTipPosAvailableCount > 1)
                {
                    // 把seqListUsedHead拆分成可用退枪头位置数目
                    /**
                     * 拆分规则：
                     * seqListUsedHead / releaseTipPosAvailableCount是否有余数，没有余数，看①；有余数，看②
                     * ①seqListUsedHead 拆分成 releaseTipPosAvailableCount 个子List<Seq>
                     * ②seqListUsedHead 拆分成 ( releaseTipPosAvailableCount - 1) 个List<Seq>，然后余数为一个新的List<Seq>
                     */
                    var splitSize = seqListUsedHead.Count / releaseTipPosAvailableCount;
                    var splitRemain = seqListUsedHead.Count % releaseTipPosAvailableCount;
                    List<List<Seq>> seqSplitList;
                    if (splitRemain == 0)
                    {
                        seqSplitList = ObjectUtils.SplitList(seqListUsedHead, splitSize);
                    }
                    else
                    {
                        // “被除数”补全到能让“除数”整除，让拆分更完整
                        splitSize = (seqListUsedHead.Count + (releaseTipPosAvailableCount - splitRemain)) / releaseTipPosAvailableCount;
                        seqSplitList = ObjectUtils.SplitList(seqListUsedHead, splitSize);
                    }

                    for (var i = 0; i < seqSplitList.Count; i++)
                    {
                        var seqList = seqSplitList.ElementAt(i);
                        foreach (var seq in seqList)
                        {
                            seq.ReleaseTipPosIndex = releaseTipPosAvailableIndexList[i];
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 执行
        /// </summary>
        /// <param name="replaceTipOverRange">超过量程分液是否换枪头</param>
        private void TakeAction(bool replaceTipOverRange)
        {
            // 每个seq占全部seqList的百分比
            double percentEachSeq = 100.0 / seqList.Count;
            // 每完成一步增加的百分比（4个基本动作：取枪头、吸液、喷液、退枪头）
            double percentEachStep = percentEachSeq / 4;
            // 每个seq运行所需时间（默认30s，通过不断计算刷新本值）
            double timeNeedEachSeq = 30;
            // 本Seq开始时间
            var dateTimeStart = DateTime.Now;
            // 本Seq结束时间
            var dateTimeStop = DateTime.Now;
            // 重复次数计数器
            var repeatTick = 0;

            for (var i = 0; i < seqList.Count; i++)
            {
                // 当前移液信息
                var seq = seqList.ElementAt(i);
                // 下一个移液信息
                var seqNext = i == seqList.Count - 1 ? null : seqList.ElementAt(i + 1);

                // 确保WAIT特殊指令能暂停
                CmdHelper.ManualStop(false);

                // 判断是否为重复执行特殊指令
                if (seq.Cmd.ToLower().Contains(AutoLiquid_Library.Utils.ConstantsUtils.RepeatCmd.ToLower()))
                {
                    // 需要重复的次数（0代表无限次）
                    var repeatTimeStr = seq.Cmd.ToLower()
                        .Replace(AutoLiquid_Library.Utils.ConstantsUtils.RepeatCmd.ToLower(), "").Trim();
                    var repeatTime = repeatTimeStr.Equals("") ? 0 : Int32.Parse(repeatTimeStr);
                    if (repeatTime != 0 && repeatTick >= repeatTime)
                    {
                        this.ControlProgressBar.SetProgressBar(100.0);
                        continue;
                    }

                    i = -1;
                    // // 初始化盘位信息
                    // Dispatcher.Invoke(InitTemplates);
                    // // 分配退枪头信息
                    // AllocateReleaseTipInfo();
                    repeatTick++;
                }

                dateTimeStart = DateTime.Now;

                // 更新倒计时
                Dispatcher.Invoke(() =>
                {
                    // 第一个Seq默认给30s倒计时
                    var totalTimeNeed = timeNeedEachSeq * (seqList.Count - i);
                    this.ControlProgressBar.UpdateTimeRemain(totalTimeNeed);
                });

                // 是否不执行该序列
                if (seq.IsComment)
                    continue;

                // 是否只有特殊指令
                if (seq.IsCmdOnly)
                {
                    // 执行特殊指令
                    ParseCmdLine(seq.Cmd, seq.IsTxtLink);
                    continue;
                }

                // 使用枪头Index
                var headUsedIndex = seq.HeadUsedIndex;
                CmdHelper.headStatusList[headUsedIndex].SeqIndex = i;

                // 枪头盒
                var tipTemplateIndex = seq.TipTemplateIndex;
                var tipTemplateAssign = seq.TipTemplateAssign;
                var tipTemplateConsumableType = seq.TipTemplateConsumableType;
                var tipChannel = seq.TipChannel;

                // 源盘
                var sourceTemplateIndex = seq.SourceTemplateIndex;
                var sourceTemplateConsumableType = seq.SourceTemplateConsumableType;
                // 20250625 SourceHoleIndex改成SourceHoleIndexList
                // var sourceHoleIndex = seq.SourceHoleIndex;
                var sourceHoleIndexList = seq.SourceHoleIndexList;
                var sourceVolumeAbsorbMore = seq.SourceVolumeAbsorbMore;

                // 靶盘
                var targetTemplateIndexList = seq.TargetTemplateIndexList;
                var targetTemplateConsumableType = seq.TargetTemplateConsumableType;
                var targetHoleIndexList = seq.TargetHoleIndexList;
                var volumeEachList = seq.VolumeEachList.Select(p => p.Calibration).ToList();

                // 梯度稀释
                var serialDilute = seq.SerialDilute;

                /**
                 * 特殊指令
                 */
                var cmd = seq.Cmd;

                // 枪头盒偏移
                var tipBoxOffsetXStr = ObjectUtils.GetCmdAccordTag(cmd,
                    AutoLiquid_Library.Utils.ConstantsUtils.TipBoxOffsetXCmd, seq.IsTxtLink);
                var tipBoxOffsetYStr = ObjectUtils.GetCmdAccordTag(cmd,
                    AutoLiquid_Library.Utils.ConstantsUtils.TipBoxOffsetYCmd, seq.IsTxtLink);
                var tipBoxOffsetZStr = ObjectUtils.GetCmdAccordTag(cmd,
                    AutoLiquid_Library.Utils.ConstantsUtils.TipBoxOffsetZCmd, seq.IsTxtLink);
                CmdHelper.tipBoxOffset.X = tipBoxOffsetXStr.Equals("") ? 0m : decimal.Parse(tipBoxOffsetXStr);
                CmdHelper.tipBoxOffset.Y = tipBoxOffsetYStr.Equals("") ? 0m : decimal.Parse(tipBoxOffsetYStr);
                CmdHelper.tipBoxOffset.Z = tipBoxOffsetZStr.Equals("") ? 0m : decimal.Parse(tipBoxOffsetZStr);

                // 退枪头盘位Index 
                var releaseTipTemplateIndexStr = ObjectUtils.GetCmdAccordTag(cmd,
                     AutoLiquid_Library.Utils.ConstantsUtils.ReleaseTipTemplateCmd, seq.IsTxtLink);
                CmdHelper.releaseTipTemplateIndex = releaseTipTemplateIndexStr.Equals("") ? (int?)null : int.Parse(releaseTipTemplateIndexStr) - 1;

                // 多吸液体返回源孔喷出
                var reJet2Source = seq.ReJet2Source;

                // 移液头最大量程
                var headLiquidRangeMax = ConstantsUtils.LiquidRangeMaxDic[ParamsHelper.HeadList[headUsedIndex].HeadLiquidRange];
                // 移液头实际量程
                var headLiquidRangeReal = Convert.ToDecimal((int)ParamsHelper.HeadList[headUsedIndex].HeadLiquidRange);

                // 靶孔数目（区分一吸一喷还是一吸多喷）
                var targetHoleIndexCount = targetHoleIndexList.Count;
                // 是否一吸多喷
                var isOne2More = targetHoleIndexCount > 1;

                /**
                 * 当前机器移液量程（x微升的机器会存在吸喷多于x微升 或者 一吸多喷多次吸液的情况，所以要采取多次吸液）
                 */
                // 分液次数
                var liquidAbsorbCount = 1;
                // 每次吸液量
                var liquidVolumeAbsorbEach = 0m;
                // 一吸多喷中，超过量程吸喷液信息
                var One2MoreInfoList = new List<One2More>();
                // 额外多吸体积
                var additionalVolume = serialDilute ? 0 : (ParamsHelper.CommonSettingList[headUsedIndex].AbsorbAirBeforePercent + ParamsHelper.CommonSettingList[headUsedIndex].AbsorbAirAfterPercent) * 0.01m * headLiquidRangeReal;

                /**
                 * 计算吸液次数和每次吸液体积
                 */
                // 一吸一喷
                if (!isOne2More)
                {
                    // 一共吸取的体积
                    var liquidVolumeTotal = 0m;
                    for (var j = 0; j < targetHoleIndexCount; j++)
                    {
                        // 如果一吸多喷投入体积只有一个，就默认为相同体积，否则体积数必须等于靶孔数
                        if (volumeEachList.Count > 1)
                            liquidVolumeAbsorbEach += volumeEachList[j];
                        else
                            liquidVolumeAbsorbEach += volumeEachList[0];
                    }
                    liquidVolumeTotal = liquidVolumeAbsorbEach;

                    // 注意：需要加上空气量再比较，（吸液前后空气量 + 吸液量）是否大于量程的最大值
                    while (liquidVolumeAbsorbEach + additionalVolume > headLiquidRangeMax)
                    {
                        liquidAbsorbCount++;
                        liquidVolumeAbsorbEach = liquidVolumeTotal / liquidAbsorbCount;
                    }
                }
                // 一吸多喷
                else
                {
                    // 上次移液截取的孔位开始index
                    var lastStartIndex = 0;
                    // lastStartIndex已经开始计数中
                    var lastStartIndexCalculating = false;

                    // 多吸体积
                    if (!serialDilute)
                    {
                        if (sourceVolumeAbsorbMore > 0)
                            additionalVolume += sourceVolumeAbsorbMore;
                        else
                        {
                            // 如果之前的seq已经多吸，则additionalVolume加上枪头液体剩余量
                            if (CmdHelper.headStatusList[headUsedIndex].VolumeAbsorbMoreLeft > 0)
                                additionalVolume += CmdHelper.headStatusList[headUsedIndex].VolumeAbsorbMoreLeft;
                            else
                                additionalVolume += ParamsHelper.CommonSettingList[headUsedIndex].AbsorbLiquidMorePercent * 0.01m * headLiquidRangeReal;
                        }
                    }

                    for (var index = 0; index < seq.TargetHoleIndexList.Count; index++)
                    {
                        // 当前孔吸液体积
                        var volumeCurrentHole = volumeEachList.Count == 1
                            ? volumeEachList[0]
                            : volumeEachList[index];
                        // 下一孔吸液体积
                        var volumeNextHole = 0m;
                        if (index + 1 < seq.TargetHoleIndexList.Count)
                        {
                            volumeNextHole = volumeEachList.Count == 1
                                ? volumeEachList[0]
                                : volumeEachList[index + 1];
                        }

                        // 当前孔吸液体积 > 量程
                        if (volumeCurrentHole + additionalVolume > headLiquidRangeMax)
                        {
                            var absorbCountInThisHoleOne2More = 1; // 本孔分液次数
                            liquidVolumeAbsorbEach = volumeCurrentHole;
                            while (liquidVolumeAbsorbEach + additionalVolume > headLiquidRangeMax)
                            {
                                absorbCountInThisHoleOne2More++;
                                liquidVolumeAbsorbEach = volumeCurrentHole / absorbCountInThisHoleOne2More;
                            }

                            for (var absortIndex = 0; absortIndex < absorbCountInThisHoleOne2More; absortIndex++)
                            {
                                var one2More = new One2More
                                {
                                    SourceVolumeAbsorbEach = liquidVolumeAbsorbEach,
                                    TargetTemplateIndexListEach = seq.TargetTemplateIndexList.Count == 1
                                    ? seq.TargetTemplateIndexList
                                    : seq.TargetTemplateIndexList.GetRange(index, 1),
                                    TargetHoleIndexListEach = seq.TargetHoleIndexList.GetRange(index, 1),
                                    TargetVolumeJetListEach = new List<decimal> { liquidVolumeAbsorbEach }
                                };
                                One2MoreInfoList.Add(one2More);

                                liquidAbsorbCount++;
                            }
                            liquidVolumeAbsorbEach = 0;
                            lastStartIndex = index;
                            lastStartIndexCalculating = false;
                        }
                        else
                        {
                            liquidVolumeAbsorbEach += volumeCurrentHole;
                            if (serialDilute) liquidVolumeAbsorbEach = volumeCurrentHole;

                            // 如果叠加体积超过量程，剔除现在的孔
                            if (liquidVolumeAbsorbEach + additionalVolume > headLiquidRangeMax && !serialDilute)
                            {
                                // 前n-1次分液（n为分液次数）
                                var indexCount = index - lastStartIndex;
                                var one2More = new One2More
                                {
                                    SourceVolumeAbsorbEach = liquidVolumeAbsorbEach - volumeCurrentHole, // 减去当前孔多余的体积
                                    TargetTemplateIndexListEach = seq.TargetTemplateIndexList.Count == 1
                                        ? seq.TargetTemplateIndexList
                                        : seq.TargetTemplateIndexList.GetRange(lastStartIndex, indexCount),
                                    TargetHoleIndexListEach =
                                        seq.TargetHoleIndexList.GetRange(lastStartIndex, indexCount),
                                    TargetVolumeJetListEach = volumeEachList.Count == 1
                                        ? volumeEachList
                                        : volumeEachList.GetRange(lastStartIndex, indexCount)
                                };
                                One2MoreInfoList.Add(one2More);

                                liquidAbsorbCount++;
                                liquidVolumeAbsorbEach = 0;
                                lastStartIndex = index;
                                lastStartIndexCalculating = false;
                                index -= 1; // 从上一个靶孔开始检索
                            }
                            else
                            {
                                if (!lastStartIndexCalculating)
                                {
                                    lastStartIndex = index;
                                    lastStartIndexCalculating = true;
                                }

                                // 如果下一孔超过量程，处理之前的孔
                                if (volumeNextHole + additionalVolume > headLiquidRangeMax && !serialDilute)
                                {
                                    var indexCount = index - lastStartIndex + 1;
                                    var one2More = new One2More
                                    {
                                        SourceVolumeAbsorbEach = liquidVolumeAbsorbEach,
                                        TargetTemplateIndexListEach = seq.TargetTemplateIndexList.Count == 1
                                            ? seq.TargetTemplateIndexList
                                            : seq.TargetTemplateIndexList.GetRange(lastStartIndex, indexCount),
                                        TargetHoleIndexListEach =
                                            seq.TargetHoleIndexList.GetRange(lastStartIndex, indexCount),
                                        TargetVolumeJetListEach = volumeEachList.Count == 1
                                            ? volumeEachList
                                            : volumeEachList.GetRange(lastStartIndex, indexCount)
                                    };
                                    One2MoreInfoList.Add(one2More);

                                    liquidAbsorbCount++;
                                    liquidVolumeAbsorbEach = 0;
                                    lastStartIndex = index;
                                    lastStartIndexCalculating = false;
                                }
                            }

                        }
                    }

                    //最后一次分液
                    if (liquidVolumeAbsorbEach != 0)
                    {
                        var lastRoundOne2More = new One2More();
                        var lastRoundIndexCount = targetHoleIndexCount - lastStartIndex;
                        lastRoundOne2More.SourceVolumeAbsorbEach = liquidVolumeAbsorbEach;
                        lastRoundOne2More.TargetTemplateIndexListEach = seq.TargetTemplateIndexList.Count == 1
                            ? seq.TargetTemplateIndexList
                            : seq.TargetTemplateIndexList.GetRange(lastStartIndex, lastRoundIndexCount);
                        lastRoundOne2More.TargetHoleIndexListEach = seq.TargetHoleIndexList.GetRange(lastStartIndex, lastRoundIndexCount);
                        lastRoundOne2More.TargetVolumeJetListEach = volumeEachList.Count == 1
                            ? volumeEachList
                            : volumeEachList.GetRange(lastStartIndex, lastRoundIndexCount);
                        One2MoreInfoList.Add(lastRoundOne2More);
                    }

                    // 再次处理一下吸液次数，预防超出Index范围
                    liquidAbsorbCount = One2MoreInfoList.Count;
                }

                /**
                 * 流程动作逻辑：
                 * 分两种情况：1吸n喷 或者 n吸1喷；n为1次~多次
                 *
                 * 情况1：1吸n喷
                 * ①如果体积只有1个且为0，且靶盘位置不为空，直接到靶盘执行喷液动作
                 * ②其余按照正常逻辑进行
                 *
                 * 情况2：n吸1喷
                 */
                // 1吸n喷
                if (sourceHoleIndexList.Count == 1)
                {
                    var sourceHoleIndex = seq.SourceHoleIndexList[0];
                    // 直接到靶盘执行喷液动作
                    if (volumeEachList.Count == 1 && volumeEachList[0] == 0 && targetTemplateIndexList.Count == 1 && targetHoleIndexList.Count == 1)
                    {
                        // 到靶盘位置喷液
                        CmdHelper.JetSingleLiquid(headUsedIndex, sourceTemplateConsumableType, sourceTemplateIndex, sourceHoleIndex, targetTemplateConsumableType,
                            seqNext?.SourceTemplateConsumableType, false,
                            targetTemplateIndexList[0], targetHoleIndexList[0], seq, liquidVolumeAbsorbEach, reJet2Source);
                    }
                    else
                    {
                        // 取枪头，第一个序列必须取枪头
                        if (seq.IsTakeTip)
                            ActionTakeTip(headUsedIndex, tipTemplateAssign, tipTemplateIndex, tipTemplateConsumableType, sourceTemplateConsumableType, tipChannel);
                        // 更新进度条
                        this.ControlProgressBar.IncreaseProgressBar(percentEachStep);

                        // 一次吸液
                        if (liquidAbsorbCount == 1)
                        {
                            // 吸液
                            CmdHelper.AbsorbLiquid(headUsedIndex, sourceTemplateConsumableType, targetTemplateConsumableType, sourceTemplateIndex, sourceHoleIndex, targetTemplateIndexList[0], seq, liquidVolumeAbsorbEach, isOne2More);
                            // 更新进度条
                            this.ControlProgressBar.IncreaseProgressBar(percentEachStep);

                            // 一吸一喷
                            if (!isOne2More)
                            {
                                CmdHelper.JetSingleLiquid(headUsedIndex, sourceTemplateConsumableType, sourceTemplateIndex, sourceHoleIndex, targetTemplateConsumableType,
                                    seqNext?.SourceTemplateConsumableType, seq.IsReleaseTip,
                                    targetTemplateIndexList[0], targetHoleIndexList[0], seq, liquidVolumeAbsorbEach, reJet2Source);
                            }
                            // 一吸多喷
                            else
                                CmdHelper.JetMultiLiquid(headUsedIndex, sourceTemplateConsumableType, sourceTemplateIndex, sourceHoleIndex, targetTemplateConsumableType, seqNext?.SourceTemplateConsumableType, seq.IsReleaseTip, targetTemplateIndexList, targetHoleIndexList, seq, volumeEachList, reJet2Source);
                            // 更新进度条
                            this.ControlProgressBar.IncreaseProgressBar(percentEachStep);
                        }
                        // 多次吸液
                        else
                        {
                            // 分液次数最后Index
                            var lastLiquidAbsorbCountIndex = liquidAbsorbCount - 1;

                            // 一吸一喷
                            if (!isOne2More)
                            {
                                for (var j = 0; j < liquidAbsorbCount; j++)
                                {
                                    // 分液超过量程是否更换枪头（本次移液序列第2个分液之后判断）
                                    if (replaceTipOverRange && j > 0)
                                        ActionTakeTip(headUsedIndex, tipTemplateAssign, tipTemplateIndex, tipTemplateConsumableType, sourceTemplateConsumableType, tipChannel);

                                    // 吸液
                                    CmdHelper.AbsorbLiquid(headUsedIndex, sourceTemplateConsumableType, targetTemplateConsumableType, sourceTemplateIndex, sourceHoleIndex, targetTemplateIndexList[0], seq,
                                        liquidVolumeAbsorbEach, false);
                                    // 更新进度条
                                    if (j == lastLiquidAbsorbCountIndex)
                                        this.ControlProgressBar.IncreaseProgressBar(percentEachStep);

                                    // 喷液
                                    // 判断喷完液后是否退枪头
                                    var releaseTip = j < lastLiquidAbsorbCountIndex ? replaceTipOverRange : seq.IsReleaseTip;
                                    CmdHelper.JetSingleLiquid(headUsedIndex, sourceTemplateConsumableType, sourceTemplateIndex, sourceHoleIndex, targetTemplateConsumableType, sourceTemplateConsumableType, releaseTip, targetTemplateIndexList[0],
                                        targetHoleIndexList[0], seq, liquidVolumeAbsorbEach, reJet2Source && lastLiquidAbsorbCountIndex == j);
                                    // 更新进度条
                                    if (j == lastLiquidAbsorbCountIndex)
                                        this.ControlProgressBar.IncreaseProgressBar(percentEachStep);

                                    // 退枪头（本次移液序列倒数第2个分液之前判断）
                                    if (replaceTipOverRange && j < lastLiquidAbsorbCountIndex)
                                        ActionReleaseTip(headUsedIndex, false, false, seq.ReleaseTipPosIndex, true);
                                }
                            }
                            // 一吸多喷
                            else
                            {
                                for (var j = 0; j < liquidAbsorbCount; j++)
                                {
                                    var one2MoreInfo = One2MoreInfoList[j];

                                    // 分液超过量程是否更换枪头（本次移液序列第2个分液之后判断）
                                    if (replaceTipOverRange && j > 0)
                                        ActionTakeTip(headUsedIndex, tipTemplateAssign, tipTemplateIndex, tipTemplateConsumableType, sourceTemplateConsumableType, tipChannel);

                                    // 吸液
                                    CmdHelper.AbsorbLiquid(headUsedIndex, sourceTemplateConsumableType, targetTemplateConsumableType, sourceTemplateIndex, sourceHoleIndex, targetTemplateIndexList[0], seq, one2MoreInfo.SourceVolumeAbsorbEach, true);
                                    // 更新进度条
                                    if (j == lastLiquidAbsorbCountIndex)
                                        this.ControlProgressBar.IncreaseProgressBar(percentEachStep);

                                    // 判断喷完液后是否退枪头
                                    var releaseTip = j < lastLiquidAbsorbCountIndex ? replaceTipOverRange : seq.IsReleaseTip;
                                    // 喷液
                                    CmdHelper.JetMultiLiquid(headUsedIndex, sourceTemplateConsumableType, sourceTemplateIndex, sourceHoleIndex, targetTemplateConsumableType, sourceTemplateConsumableType, releaseTip, one2MoreInfo.TargetTemplateIndexListEach,
                                        one2MoreInfo.TargetHoleIndexListEach, seq, one2MoreInfo.TargetVolumeJetListEach, reJet2Source && lastLiquidAbsorbCountIndex == j);
                                    // 更新进度条
                                    if (j == lastLiquidAbsorbCountIndex)
                                        this.ControlProgressBar.IncreaseProgressBar(percentEachStep);

                                    // 退枪头（本次移液序列倒数第2个分液之前判断）
                                    if (replaceTipOverRange && j < lastLiquidAbsorbCountIndex)
                                        ActionReleaseTip(headUsedIndex, false, false, seq.ReleaseTipPosIndex, true);
                                }
                            }
                        }
                    }
                }
                // n吸1喷
                else if (sourceHoleIndexList.Count > 1)
                {
                    // 取枪头，第一个序列必须取枪头
                    if (seq.IsTakeTip)
                        ActionTakeTip(headUsedIndex, tipTemplateAssign, tipTemplateIndex, tipTemplateConsumableType, sourceTemplateConsumableType, tipChannel);
                    // 更新进度条
                    this.ControlProgressBar.IncreaseProgressBar(percentEachStep);

                    // 吸液
                    foreach (var sourceHoleIndex in sourceHoleIndexList)
                    {
                        CmdHelper.AbsorbLiquid(headUsedIndex, sourceTemplateConsumableType, targetTemplateConsumableType, sourceTemplateIndex, sourceHoleIndex, targetTemplateIndexList[0], seq,
                            liquidVolumeAbsorbEach, false);
                        // 更新进度条
                        this.ControlProgressBar.IncreaseProgressBar(percentEachStep);
                    }
                    // 喷液
                    CmdHelper.JetSingleLiquid(headUsedIndex, sourceTemplateConsumableType, sourceTemplateIndex, null, targetTemplateConsumableType,
                        seqNext?.SourceTemplateConsumableType, seq.IsReleaseTip,
                        targetTemplateIndexList[0], targetHoleIndexList[0], seq, liquidVolumeAbsorbEach, false, true);
                    // 更新进度条
                    this.ControlProgressBar.IncreaseProgressBar(percentEachStep);
                }

                // 执行特殊指令
                ParseCmdLine(cmd, seq.IsTxtLink);

                /**
                 * 退枪头
                 */
                if (seq.IsReleaseTip && CmdHelper.headStatusList[headUsedIndex].Head != EHeadStatus.TipReleased)
                    ActionReleaseTip(headUsedIndex, seqNext != null && seqNext.HeadUsedIndex != headUsedIndex, false, seq.ReleaseTipPosIndex, true);

                // 更新进度条
                if (seqNext == null) // 最后的seq
                    this.ControlProgressBar.SetProgressBar(100.0);
                else
                    this.ControlProgressBar.SetProgressBar(percentEachSeq * (i + 1));

                // 计算本次Seq所用秒数
                dateTimeStop = DateTime.Now;
                timeNeedEachSeq = (dateTimeStop - dateTimeStart).TotalSeconds;
            }
        }

        /// <summary>
        /// 取枪头 
        /// </summary>
        /// <param name="headUsedIndex">使用的移液头Index</param>
        /// <param name="tipTemplateAssign">是否指定盘位取枪头</param>
        /// <param name="tipTemplateIndex">取枪头盘位</param>
        /// <param name="consumableType">耗材类型</param>
        /// <param name="consumableTypeNext">下一个耗材类型</param>
        /// <param name="tipChannel2DArray">取枪头数量二维数组</param>
        private void ActionTakeTip(int headUsedIndex, bool tipTemplateAssign, int tipTemplateIndex, Consumable consumableType, Consumable consumableTypeNext, int[,] tipChannel2DArray)
        {
            // 全部枪头盘位
            var allTipTemplateList = tipTemplateAssign ? tipTemplateDict.Values.Where(p => p.HeadUsedIndex == headUsedIndex && p.TipBoxTemplateIndex == tipTemplateIndex).ToList() : tipTemplateDict.Values.Where(p => p.HeadUsedIndex == headUsedIndex).ToList();
            // 全部枪头盘位Index
            var allTipTemplateIndexList = tipTemplateAssign ? tipTemplateDict.Where(p => p.Value.HeadUsedIndex == headUsedIndex && p.Value.TipBoxTemplateIndex == tipTemplateIndex).Select(p => p.Key).ToList() : tipTemplateDict.Where(p => p.Value.HeadUsedIndex == headUsedIndex).Select(p => p.Key).ToList();
            // 本次所用的枪头行、列index列表
            var tipUsedIndexThisTimeList = new List<RowCol>();
            // 用到的枪头盘位Index
            var tipTemplateUsedIndex = tipTemplateIndex;
            // 取枪头开始Index
            var tipTakeStartIndex = new HoleIndex();
            // 是否有足够的枪头数
            var isTipCountEnough = false;
            // 下一个取枪头位置
            var nextTipIndex = 0;

            // 计算取枪头信息
            nextTipIndex = TipHelper.CalcTipPos(headUsedIndex, consumableType, tipChannel2DArray, allTipTemplateList, ref tipTemplateUsedIndex, ref tipUsedIndexThisTimeList, ref tipTakeStartIndex, ref isTipCountEnough);

            // 判断枪头是否足够，并提示
            if (!isTipCountEnough)
            {
                // 先复位
                CmdHelper.InitMachineAndEasy2Put(true, false);
                Dispatcher.Invoke(() =>
                {
                    // 需要置满的枪头盘位
                    var tipTemplateIndexNeedFillList = allTipTemplateIndexList.Select(p => p + 1).ToList();
                    var tipTemplateIndexNeedFillStr = String.Join("，", tipTemplateIndexNeedFillList);
                    // 提示枪头盘位置满枪头盒
                    if (MessageBox.Show((string)this.FindResource("Template") + " " + tipTemplateIndexNeedFillStr + " " + (string)this.FindResource("Prompt_Tip_Not_Enough_Pls_Fill_And_Continue"),
                            (string)this.FindResource("Prompt"), MessageBoxButton.OK, MessageBoxImage.Warning) ==
                        MessageBoxResult.OK)
                    {
                        foreach (var tipTemplate in allTipTemplateList)
                        {
                            // 如果设置的开始位置与原来的一样，,必须先设为-1，再设置
                            if (tipTemplate.SplitButtonTipsBoxPos.SelectedIndex == 0)
                            {
                                tipTemplate.SplitButtonTipsBoxPos.SelectedIndex = -1;
                                tipTemplate.SplitButtonTipsBoxPos.SelectedIndex = 0;
                            }
                            else
                                tipTemplate.SplitButtonTipsBoxPos.SelectedIndex = 0;
                        }

                        // 重新计算取枪头信息
                        nextTipIndex = TipHelper.CalcTipPos(headUsedIndex, consumableType, tipChannel2DArray, allTipTemplateList, ref tipTemplateUsedIndex, ref tipUsedIndexThisTimeList, ref tipTakeStartIndex, ref isTipCountEnough);
                    }
                });
            }

            // 取枪头
            CmdHelper.TakeTipInfo.TipUsedIndexList = tipUsedIndexThisTimeList;
            CmdHelper.TakeTip(headUsedIndex, consumableType, consumableTypeNext, tipTemplateUsedIndex, tipTakeStartIndex);

            // 枪头盒灵活取枪头
            var tipTemplateUsed = tipTemplateDict[tipTemplateUsedIndex];
            // 多通道移液头间距与枪头盒孔距比例
            var multiChannelStepAndTipBoxStepRelation = TipHelper.MultiChannelStepAndConsumableStepRelation(headUsedIndex, tipChannel2DArray, consumableType);
            // 枪头盒灵活取枪头 或者 多通道移液头间距与枪头盒孔距比例不一致
            if (tipTemplateUsed.TipBoxFlexible || multiChannelStepAndTipBoxStepRelation != 1.0m)
            {
                foreach (var rowCol in tipUsedIndexThisTimeList)
                {
                    tipTemplateUsed.TipBoxUsedStatus2DArray[rowCol.Row, rowCol.Col] = true;
                }
                Dispatcher.Invoke(() =>
                {
                    tipTemplateUsed.RefreshTemplateHolesColor();
                });
            }
            else
            {
                // 更新上一个孔颜色和下一个孔位置（直接更改SelectedIndex就可以）
                Dispatcher.Invoke(() =>
                {
                    if (nextTipIndex < 0)
                    {
                        tipTemplateUsed.SplitButtonTipsBoxPos.SelectedIndex = nextTipIndex;
                    }
                    else
                    {
                        // 把nextTipIndex转为孔名称（如A1、B2等），避免枪头盒下拉列表总数和nextTipIndex不是对应关系造成的取枪头定位问题
                        var posStr = ConsumableHelper.GetHolePosStr(nextTipIndex, consumableType.RowCount, consumableType.ColCount, ParamsHelper.CommonSettingList[headUsedIndex].A1Pos);
                        tipTemplateUsed.SplitButtonTipsBoxPos.SelectedItem = posStr;
                        // tipTemplateUsed.SplitButtonTipsBoxPos.SelectedIndex = nextTipIndex;
                    }
                });
            }
        }

        /// <summary>
        /// 退枪头
        /// </summary>
        /// <param name="headUsedIndex">使用的移液头Index</param>
        /// <param name="isNeedZ0AfterRelease">退枪头后是否需要高度复位</param>
        /// <param name="isMidway">是否中途退枪头</param>
        /// <param name="releaseTipPosIndex">退枪头位置Index</param>
        /// <param name="isNeedManualStop"></param>
        private void ActionReleaseTip(int headUsedIndex, bool isNeedZ0AfterRelease, bool isMidway, int releaseTipPosIndex, bool isNeedManualStop)
        {
            CmdHelper.ReleaseTip(headUsedIndex, isNeedZ0AfterRelease, isMidway, releaseTipPosIndex, isNeedManualStop);

            // TODO 如果枪头退回到取枪头位置，就把取枪头位置自动置满，可参考凌恩25盘位做法
            // Dispatcher.Invoke(() =>
            // {
            //     if (releaseTip2Origin)
            //     {
            //         var tipTemplateUsed = tipTemplateDict[CmdHelper.TakeTipInfo.TemplateIndex];
            //         // 全部枪头置满
            //         if (CmdHelper.TakeTipInfo.TipUsedIndexList.Count == tipTemplateUsed.TipBoxUsedStatus2DArray.Length)
            //         {
            //             Dispatcher.Invoke(() =>
            //             {
            //                 tipTemplateUsed.SplitButtonTipsBoxPos.SelectedIndex = 0;
            //             });
            //         }
            //         // 部分枪头置满
            //         else
            //         {
            //             foreach (var rowCol in CmdHelper.TakeTipInfo.TipUsedIndexList)
            //             {
            //                 tipTemplateUsed.TipBoxUsedStatus2DArray[rowCol.Row, rowCol.Col] = false;
            //             }
            //             Dispatcher.Invoke(() =>
            //             {
            //                 tipTemplateUsed.RefreshTemplateHolesColor();
            //             });
            //         }
            //     }
            // });
        }

        /// <summary>
        /// 解析特殊指令
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="isTxtLink">是否文件链接</param>
        public void ParseCmdLine(string cmd, bool isTxtLink)
        {
            // 拆分行
            string[] lineCmd = cmd.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lineCmd)
            {
                // 剔除相关特殊指令
                if (isTxtLink && ObjectUtils.CheckCmdExist(line))
                    continue;

                // 拆分分号
                string[] semiCmd = line.Split(new[] { ';', '；' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var semi in semiCmd)
                {
                    // 剔除相关特殊指令
                    if (!isTxtLink && ObjectUtils.CheckCmdExist(semi))
                        continue;

                    // 等待指令
                    if (semi.ToLower().StartsWith("wait"))
                    {
                        // 剔除“WAIT”
                        var subCmd = semi.Substring(4).Trim();

                        // 一直暂停
                        if (subCmd.Equals(""))
                        {
                            Dispatcher.Invoke(() =>
                            {
                                // 马上停止下一个指令执行
                                CmdHelper.isManualPause = true;
                                CmdHelper.isManualStop = false;
                                // 模拟手动点击“暂停”
                                SetLoadingStatus(ERunStatus.Pause, false);
                            });
                        }
                        // 暂停，并弹出框提示
                        else if (subCmd.Contains("*"))
                        {
                            Dispatcher.Invoke(() =>
                            {
                                // 马上停止下一个指令执行
                                CmdHelper.isManualPause = true;
                                CmdHelper.isManualStop = false;
                                // 模拟手动点击“暂停”
                                SetLoadingStatus(ERunStatus.Pause, false);
                                this.LabelPrompt.Content = subCmd.Substring(1);
                            });
                        }
                        // 暂停x秒
                        else
                        {
                            var totalSec = Utils.DataHelper.GetTotalSeconds(subCmd);
                            ActionWait(totalSec, (string)this.FindResource("Wait"));
                        }
                    }
                    // 普通指令
                    else
                    {
                        // 检测试管摆放指令
                        if (semi.ToLower().Contains(AutoLiquid_Library.Utils.ConstantsUtils.CheckTubeCmd.ToLower()))
                        {
                            var otherCmd = semi.ToLower().Replace(AutoLiquid_Library.Utils.ConstantsUtils.CheckTubeCmd.ToLower(), "");
                            CheckTube();

                            CmdHelper.DoCmd(otherCmd.Trim(), true);
                        }
                        else
                            CmdHelper.DoCmd(semi.Trim(), true);
                    }
                }
            }
        }

        /// <summary>
        /// 检测试管摆放（如没摆放，就等待）
        /// </summary>
        private void CheckTube()
        {
            // 是否检测完成
            var isCheckSuccess = false;
            CmdHelper.frmDAE.IsTubeExist = false;
            // 检测次数
            var trueCheckCount = 3;
            // 连续为true的次数
            var trueContinuousCount = 0;

            /**
             * 判断逻辑：
             * ①快速检测：快速检测10次（检测间隔为20ms），不弹出提示框，如果连续3次true，则继续后面的步骤； 如果false，跳到②普通检测
             * ②普通检测：弹出提示屏幕，检测3次（检测间隔为500ms）
             */
            // 快速检测
            for (var i = 0; i < 10; i++)
            {
                // 发送检测指令
                CmdHelper.CheckTubeExist(false);
                Thread.Sleep(20);
                CmdHelper.ManualStop(false);

                if (CmdHelper.frmDAE.IsTubeExist)
                {
                    trueContinuousCount++;
                    if (trueContinuousCount >= trueCheckCount)
                    {
                        isCheckSuccess = true;
                        break;
                    }
                }
                else
                    trueContinuousCount = 0; // 只要有1次不为true，就复位
            }

            // 普通检测
            if (!isCheckSuccess)
            {
                trueContinuousCount = 0;

                // 弹出框提示
                WindowMessageBox wmb = null;
                Dispatcher.Invoke(() =>
                {
                    wmb = new WindowMessageBox();
                    wmb.Show();
                });

                while (trueContinuousCount < trueCheckCount)
                {
                    // 发送检测指令
                    CmdHelper.CheckTubeExist(false);
                    Thread.Sleep(500);
                    CmdHelper.ManualStop(false);

                    if (CmdHelper.frmDAE.IsTubeExist)
                        trueContinuousCount++;
                    else
                        trueContinuousCount = 0; // 只要有1次不为true，就复位
                }

                Dispatcher.Invoke(() =>
                {
                    wmb?.Close();
                });
            }

            // 复位属性
            CmdHelper.frmDAE.IsTubeExist = false;
        }

        /// <summary>
        /// 时间等待
        /// </summary>
        /// <param name="waitSec">时间</param>
        /// <param name="title">显示标题</param>
        private void ActionWait(int waitSec, string title)
        {
            LogHelper.Info(title, waitSec + (string)this.FindResource("Second"));
            Dispatcher.Invoke(() =>
            {
                SetLoadingStatus(ERunStatus.Countdown, false);

                this.LabelIntervalRemainTitle.Content = title + (string)this.FindResource("IntervalRemain");
                this.LabelIntervalRemain.Content = waitSec;

                this.StackPanelIntervalRemain.Visibility = Visibility.Visible;
            });
            // 已经过去的时间
            var waitTimePassed = 0;
            while (waitTimePassed < waitSec)
            {
                Thread.Sleep(1000);
                waitTimePassed += 1;

                Dispatcher.Invoke(() =>
                {
                    var waitRemain = waitSec - waitTimePassed;
                    this.LabelIntervalRemain.Content = waitRemain;
                });
                CmdHelper.ManualStop(false);
            }

            Dispatcher.Invoke(() =>
            {
                this.StackPanelIntervalRemain.Visibility = Visibility.Collapsed;
                SetLoadingStatus(ERunStatus.Running, false);
            });
        }
    }
}
