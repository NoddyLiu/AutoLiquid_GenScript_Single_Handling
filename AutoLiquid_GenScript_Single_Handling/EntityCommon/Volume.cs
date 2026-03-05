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
    /// 液体体积 model
    /// </summary>
    public class Volume
    {
        // 原始体积（一般用于界面显示）
        public decimal Original = 0m;

        // 校准体积（真实的移液体积，不论有没有启用校准功能）
        public decimal Calibration = 0m;
    }
}
