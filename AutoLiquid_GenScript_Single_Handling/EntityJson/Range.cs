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
    /// 最大量程设置
    /// </summary>
    [Serializable()]
    public class Range
    {
        // 10微升
        public decimal Ten = 15m;

        // 20微升
        public decimal Twenty = 20m;

        // 200微升
        public decimal TwoHundred = 200m;

        // 1000微升
        public decimal OneThousand = 1000m;

        // 5000微升
        public decimal FiveThousand = 5000m;
    }
}
