using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace AutoLiquid_GenScript_Single_Handling.EntityJson
{
    /// <summary>
    /// 标准机原点坐标
    /// </summary>
    [Serializable()]
    public class OriginalPoint
    {
        // 盘位标准值
        public Position PosTemplate = new Position();

        // 预退枪头标准值
        public Position PosPreReleaseTip = new Position();

        // 退枪头标准值
        public Position PosReleaseTip = new Position();
    }
}
