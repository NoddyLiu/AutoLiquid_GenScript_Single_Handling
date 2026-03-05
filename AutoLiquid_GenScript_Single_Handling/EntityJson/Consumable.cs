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
    /// 耗材
    /// </summary>
    [Serializable()]
    public class Consumable
    {
        /*****************************************************公用属性start************************************************************/
        // 组名（根据组名与Excel中的耗材类型匹配）
        public string GroupName = "";

        // 是否为枪头盒
        public bool IsTipBox = false;

        /**
         * 行走高度
         */
        // 行走高度
        public decimal NormalHeight = 0.00m;
        // 跨耗材高度复位（两个耗材之间行走如果有碰撞，需要置为true）
        public bool NormalHeightReset = false;

        /**
         * 盘位坐标
         */
        // 首孔坐标
        public List<Position> HoleStartPosList = new List<Position>();
        // 首孔坐标是否可用
        public List<bool> TemplateAvailableList = new List<bool>();
        // 自动填充
        public bool HoleStartPosAutoFill = false;
        // 盘位行列距
        public Position TemplateStep = new Position { X = 0.00m, Y = 0.00m };
        // 占用盘位个数
        public ESpan TemplateOccupySpan = ESpan.One;

        /**
         * 孔行列设置
         */
        // 行数
        public int RowCount = 8;
        // 列数
        public int ColCount = 12;
        // 孔距
        public Position HoleStep = new Position { X = 9.00m, Y = 9.00m };

        /**
         * 可变距设置
         */
        // 变距步数
        public int VariableDistanceStep = 0;
        /*****************************************************公用属性end************************************************************/



        /*****************************************************枪头盒特有属性start************************************************************/
        /**
         * 取枪头设定
         */
        // 预取枪头高度
        public decimal PrepareTakeTipHeight = 0.00m;
        // 取枪头高度
        public decimal TakeTipHeight = 0.00m;
        // 取枪头速度
        public string TakeTipSpeedCmd = "";
        // 重复取枪头次数
        public int TakeTipRepeatTime = 0;
        // 重复取枪头高度
        public decimal TakeTipRepeatHeight = 0.00m;
        // 取枪头后预提高度
        public decimal TakeTipAfterPrepareHeight = 0.00m;
        // 预提高度后执行指令
        public string TakeTipAfterPrepareHeightCmd = "";
        /*****************************************************枪头盒特有属性end************************************************************/



        /*****************************************************一般耗材特有属性start************************************************************/
        /**
         * 吸液设定
         */
        // 吸液高度
        public decimal LiquidAbsorbHeight = 0.00m;
        // 吸液速度
        public string AbsorbSpeed = "";
        // 吸液后等待
        public decimal LiquidAbsorbDelay = 0.00m;
        // 吸液提起速度
        public string AbsorbHeight2NormalHeightSpeed = "";
        // 吸液提起高度
        public decimal AbsorbHeight2LiftingHeight = 0.00m;
        // 吸液提起后等待
        public decimal LiquidAbsorbDelayAfterLift = 0.00m;

        /**
         * 喷液设定
         */
        // 喷液高度
        public decimal LiquidJetHeight = 0.00m;
        // 喷液速度
        public string JetSpeed = "";
        // 喷液后等待
        public decimal LiquidJetDelay = 0.00m;
        // 喷液提起速度
        public string JetHeight2NormalHeightSpeed = "";

        /**
         * 靠壁设定
         */
        // 靠壁高度
        public decimal LiquidJetWallHeight = 0.00m;
        // 靠壁偏移
        public decimal LiquidJetWallOffset = 0m;
        // 靠壁触发条件
        public decimal LiquidJetWallTrigger = 0m;
        // 靠壁喷液
        public bool WallJet = false;

        /**
         * 混合设定
         */
        // 吸前混合高度
        public decimal AbsorbMixingHeight = 0.00m;
        // 吸前混合速度
        public string AbsorbMixingSpeed = "";
        // 喷后混合高度
        public decimal JetMixingHeight = 0.00m;
        // 喷后混合速度
        public string JetMixingSpeed = "";
        /*****************************************************一般耗材特有属性end************************************************************/
    }
}
