using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace AutoLiquid_GenScript_Single_Handling.EntityJson
{
    /// <summary>
    /// 机器偏移值（本机值 - 标准机值）
    /// </summary>
    [Serializable()]
    public class Device
    {
        // 机器id
        public string Id  = "";

        // 盘位偏移值
        public Position OffsetTemplate  = new Position();

        // 退枪头偏移值
        public Position OffsetReleaseTip  = new Position();
    }
}
