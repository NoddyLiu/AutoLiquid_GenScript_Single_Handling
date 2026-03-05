using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoLiquid_Library.Enum;

namespace AutoLiquid_GenScript_Single_Handling.Utils
{
    public class ConstantsUtils
    {
        // 量程最大值
        public static Dictionary<ELiquidRange, decimal> LiquidRangeMaxDic = new Dictionary<ELiquidRange, decimal>
        {
            {ELiquidRange.Ten, ParamsHelper.Range.Ten},
            {ELiquidRange.Twenty, ParamsHelper.Range.Twenty },
            {ELiquidRange.TwoHundred, ParamsHelper.Range.TwoHundred},
            {ELiquidRange.OneThousand, ParamsHelper.Range.OneThousand},
            {ELiquidRange.FiveThousand, ParamsHelper.Range.FiveThousand},
        };

        // 主界面每个盘位占据Grid的多少比重
        public static int TemplateOccupyGridSpan = 2;
    }
}
