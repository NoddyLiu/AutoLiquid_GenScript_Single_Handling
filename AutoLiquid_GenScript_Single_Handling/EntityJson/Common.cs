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
    /// 移液头通用设置
    /// </summary>
    [Serializable()]
    public class Common
    {
        /**
         * 取枪头
         */
        // 逐列取枪头
        public bool TakeTipEachCol = false;
        // 取枪头方向（默认从左往右取，false则为从右往左取）
        public bool TakeTipLeft2Right = true;

        /**
         * 退枪头
         */
        // 预退枪头位置
        public List<Position> PrepareReleaseTipPosList = new List<Position>();
        // 退枪头位置
        public List<Position> ReleaseTipPosList = new List<Position>();
        // 退枪头变距（步数）
        public int ReleaseTipVariableDistanceStep = 0;
        // 退枪头位置启用情况
        public List<bool> ReleaseTipPosAvailableList = new List<bool>();

        // 预退枪头先走X轴
        public bool PrepareReleaseTipAxisXGoFirst = false;
        // 预退枪头先走Y轴
        public bool PrepareReleaseTipAxisYGoFirst = false;
        // 推脱板偏移
        public decimal ReleaseTipOffset = 0;
        // 退枪头后指令
        public string ReleaseTipAfterCmd = "";
        // 返回取枪头位置退枪头
        public bool ReleaseTipBack2TakePos = false;
        // 退枪头前高度回零
        public bool ReleaseTipZa0Before = false;
        // 退枪头后高度回零
        public bool ReleaseTipZa0After = false;
        // 退枪头速度指令
        public string ReleaseTipSpeedCmd = "";

        /**
         * 速度设定
         */
        // 默认X
        public string DefaultXSpeed = "";
        // X速度百分比
        public decimal XSpeedPercent = 100.00m;
        // 默认Y
        public string DefaultYSpeed = "";
        // Y速度百分比
        public decimal YSpeedPercent = 100.00m;
        // 默认Z
        public string DefaultZSpeed = "";
        // Z速度百分比
        public decimal ZSpeedPercent = 100.00m;
        // 默认P
        public string DefaultPSpeed = "";
        // P速度百分比
        public decimal PSpeedPercent = 100.00m;

        /**
         * 耗材（可添加删除）
         */
        public List<Consumable> Consumables = new List<Consumable>();

        /**
         * 吸喷液前处理
         */
        // 吸液前吸空气量（量程百分比）
        public decimal AbsorbAirBeforePercent = 0.0m;
        // 喷液后喷空气等待时间（S）
        public decimal AirDelayAfterJet = 0.0m;
        // 吸液后吸空气量（量程百分比）
        public decimal AbsorbAirAfterPercent = 0.0m;
        // 一吸多喷吸液后多吸体积（量程百分比）
        [Obsolete("弃用，请使用属性AbsorbLiquidMore", true)]
        public decimal AbsorbLiquidMoreOne2MorePercent = 0.0m;
        // 一吸多喷吸液后喷出体积比例（多吸体积的百分比）
        [Obsolete("弃用，请使用属性ReverseJetAfterAbsorb", true)]
        public decimal JetLiquidMoreOne2MoreScale = 0.0m;
        // 多吸体积（量程百分比）（一吸多喷专用）
        public decimal AbsorbLiquidMorePercent = 0.0m;
        // 吸后反喷体积（量程百分比，即先吸入该属性的体积，再全部喷出该属性体积）
        public decimal ReverseJetAfterAbsorbPercent = 0.0m;
        // 多点校准（用于校准喷液轴）
        public Calibration MultiCalibration = new Calibration();

        /**
         * 分段喷液
         */
        // 第1段喷液速度
        public string SpeedJet1 = "";
        // 第1段喷液速度百分比
        public decimal Jet1SpeedPercent = 100.00m;
        // 第2段喷液速度
        public string SpeedJet2 = "";
        // 第2段喷液速度百分比
        public decimal Jet2SpeedPercent = 100.00m;
        // 第2段喷液体积
        public decimal VolumeJet2 = 0.0m;
        // 两段喷液间停留时间（ms）
        public int DelayBetweenJet = 0;

        /**
         * 枪头分段提起
         */
        // 第1段提起速度
        public string SpeedTipLift1 = "";
        // 第1段提起速度百分比
        public decimal TipLift1SpeedPercent = 100.00m;
        // 第1段提起高度
        public decimal TipLift1Height = 0.00m;

        /**
         * 推出盘位
         */
        public Position PositionLaunchPlate = new Position();

        /**
         * 盘位摆放（即A1方向）
         */
        public EA1Pos A1Pos = EA1Pos.LeftTop;
    }
}
