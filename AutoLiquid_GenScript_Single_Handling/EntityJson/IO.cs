using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using AutoLiquid_Library.Enum;

namespace AutoLiquid_GenScript_Single_Handling.EntityJson
{
    /// <summary>
    /// 输入输出
    /// </summary>
    [Serializable()]
    public class IO
    {
        // Tcp连接
        public bool Tcp = true;

        // Ip地址
        public string IP = "192.168.1.3";

        // 是否启用仓门
        public bool DoorAvailable = false;
        // 仓门查询指令
        public string DoorCmd = "AD2";
        // 仓门停止功能打开指令
        public string DoorAvailableCmd = "A9 1";
        // 是否启用急停
        public bool EmergencyStopAvailable = false;
        // 急停查询指令
        public string EmergencyStopCmd = "AD1";
        // 急停功能打开指令
        public string EmergencyStopAvailableCmd = "A9 2";


        /**
         * 外部设备通信
         */
        // 是否打开串口通信
        public bool SerialPortAvailable = false;
        // 串口号
        public string SeialPort = "COM1";

        /**
         * 设备情况
         */
        /**
         * 风扇
         */
        // 是否启用风扇
        public bool FanAvailable = false;
        // 风扇开指令
        public string FanOpenCmd = "AO21";
        // 风扇关指令
        public string FanCloseCmd = "AO20";
        /**
         * 紫外灯
         */
        // 是否启用紫外灯
        public bool UVAvailable = false;
        // 紫外灯开指令
        public string UVOpenCmd = "AO11";
        // 紫外灯关指令
        public string UVCloseCmd = "AO10";
        /**
         * 照明灯
         */
        // 是否启用照明
        public bool LightAvailable = false;
        // 照明开指令
        public string LightOpenCmd = "AO31";
        // 照明关指令
        public string LightCloseCmd = "AO30";
        /**
         * 警示灯
         */
        // 是否启用警示灯
        public bool WarningLightAvailable = false;
        // 警示灯绿色开指令
        public string WarningLightGreenOpenCmd = "AO11";
        // 警示灯黄色开指令
        public string WarningLightYellowOpenCmd = "AO21";
        // 警示灯红色开指令
        public string WarningLightRedOpenCmd = "AO31";
    }
}
