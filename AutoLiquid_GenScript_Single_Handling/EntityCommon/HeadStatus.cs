using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Shapes;
using AutoLiquid_GenScript_Single_Handling.EntityJson;
using AutoLiquid_Library.Enum;

namespace AutoLiquid_GenScript_Single_Handling.EntityCommon
{
    /// <summary>
    /// 移液头状态信息
    /// </summary>
    public class HeadStatus
    {
        // 执行到哪一步
        public EHeadStatus Head = EHeadStatus.TipReleased;

        // 执行到哪一个Seq
        public int SeqIndex = 0;

        // 枪头当前剩余多吸体积
        public decimal VolumeAbsorbMoreLeft = 0m;
        // 枪头当前剩余空气体积
        public decimal VolumeAirLeft = 0m;

        /**
         * 源盘信息
         */
        // 耗材
        public Consumable SourceConsumableType;
        // 盘位Index
        public int SourceTemplateIndex;
        // 孔位Index
        public HoleIndex SourceHoleIndex;
    }
}
