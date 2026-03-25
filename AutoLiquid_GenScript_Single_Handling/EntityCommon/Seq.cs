using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoLiquid_GenScript_Single_Handling.EntityJson;
using AutoLiquid_Library.Enum;

namespace AutoLiquid_GenScript_Single_Handling.EntityCommon
{
    /// <summary>
    /// 序列（枪头位置 -> 源盘位置 -> 靶盘位置）
    /// </summary>
    public class Seq
    {
        /**
         * 取枪头
         */
        // 枪头盘位Index
        public int TipTemplateIndex;
        // 是否指定盘位取枪头（默认为否：枪头会按顺序逐个枪头盒取；是：到指定盘位取枪头）
        public bool TipTemplateAssign = false;
        // 枪头盘耗材类型
        public Consumable TipTemplateConsumableType;
        // 使用枪头数目[行, 列]
        public int[,] TipChannel = new int[1, 1];
        // 是否取枪头
        public bool IsTakeTip = true;

        /**
         * 源盘
         */
        // 源盘盘名
        public string SourceTemplateName = "";
        // 源盘盘位Index
        public int SourceTemplateIndex;
        // // 源盘孔Index
        // public HoleIndex SourceHoleIndex = new HoleIndex();
        // 源盘孔Index （孔数为1个：1次吸液，n次喷液；孔数为多个：n次吸液，1次喷液。n为1到多）
        public List<HoleIndex> SourceHoleIndexList = new List<HoleIndex>();
        // 源盘耗材类型
        public Consumable SourceTemplateConsumableType;
        // 吸液后多吸体积（微升，也可在参数界面设置量程百分比）
        public decimal SourceVolumeAbsorbMore = 0;
        // 吸后反喷体积（微升，也可在参数界面设置量程百分比）
        public decimal SourceVolumeReverseJet = 0;
        // 多吸液体返回源孔喷出
        public bool ReJet2Source = false;

        /**
         * 靶盘（可能存在一吸多喷）
         */
        // 靶盘盘名
        public string TargetTemplateName = "";
        // 靶盘盘位Index
        public List<int> TargetTemplateIndexList = new List<int>();
        // 靶盘孔Index
        public List<HoleIndex> TargetHoleIndexList = new List<HoleIndex>();
        // 靶盘耗材类型
        public Consumable TargetTemplateConsumableType;
        // 是否梯度稀释
        public bool SerialDilute;
        // 喷液体积补偿
        public List<JetOffset> JetOffsetList = new List<JetOffset>();
        // 一吸多喷喷液后回吸体积
        public decimal VolumeBackAbsorb = 0;

        /**
         * 体积
         */
        // 每次体积
        public List<Volume> VolumeEachList = new List<Volume>();

        /**
         * 移液轮次
         * 一轮最多n块板位置（n以盘位数为准，包括枪头盒、源盘、靶盘），超过n块，轮次加1，以此类推
         */
        public int Round = 1;

        /**
         * 混合信息
         */
        // 吸前混合体积
        public decimal AbsorbMixingVolume;
        // 吸前混合次数
        public int AbsorbMixingCount;
        // 喷后混合体积
        public decimal JetMixingVolume;
        // 喷后混合次数
        public int JetMixingCount;

        /**
         * 吸液靠壁信息
         */
        // 吸液靠壁
        public List<EWall> AbsorbWallList = new List<EWall>();
        // 喷液靠壁
        public List<EWall> JetWallList = new List<EWall>();

        /**
           * 特殊指令
           */
        // 是否文件链接
        public bool IsTxtLink = false;
        // 是否只有特殊指令
        public bool IsCmdOnly;
        public string Cmd = "";
        public string CmdAbsorbBefore = "";
        public string CmdAbsorbAfter = "";
        public string CmdJetBefore = "";
        public string CmdJetAfter = "";
        public string CmdAbsorbMixingBefore = "";
        public string CmdAbsorbMixingAfter = "";
        public string CmdJetMixingBefore = "";
        public string CmdJetMixingAfter = "";
        // 是否注释该行序列（即不执行）
        public bool IsComment;

        // 使用的移液头Index
        public int HeadUsedIndex = 0;

        /**
         * 退枪头
         */
        // 是否退枪头
        public bool IsReleaseTip = true;
        // 退枪头位置
        public int ReleaseTipPosIndex = 0;

        // 是否泵分液
        public bool IsPumpLiquid = false;

        // ── EP管架引物名称（A列，每行对应源盘或靶盘某个孔位的EP管引物名称）──
        // 当源盘或靶盘为EP管架（GroupName含"ep"）时，此字段存放该行对应的引物名称（primerlabel）
        // 扫码时作为该孔EP管的期望条码
        public string EpTubePrimerLabel = "";
    }
}
