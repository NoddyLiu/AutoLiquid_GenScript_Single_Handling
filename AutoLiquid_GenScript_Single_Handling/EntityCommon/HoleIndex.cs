using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Shapes;

namespace AutoLiquid_GenScript_Single_Handling.EntityCommon
{
    /// <summary>
    /// 孔位Index
    /// </summary>
    public class HoleIndex
    {
        // 原始位置Index
        public int OriIndex = 0;

        /**
         * 以下用于灵活取枪头、吸喷液
         */
        // X轴孔偏移个数（正数为正向偏移，负数为负向偏移）
        public int XHoleOffset = 0;
        // 移液头通道X间距与耗材X间距是否不一致（如果不一致，先走移液头的偏移，再走耗材的偏移）
        public bool StepNotSameX = false;
        // Y轴孔偏移个数（正数为正向偏移，负数为负向偏移）
        public int YHoleOffset = 0;
        // 移液头通道Y间距与耗材Y间距是否不一致（如果不一致，先走移液头的偏移，再走耗材的偏移）
        public bool StepNotSameY = false;
    }
}
