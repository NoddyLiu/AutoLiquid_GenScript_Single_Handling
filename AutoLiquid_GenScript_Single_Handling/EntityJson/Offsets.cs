using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace AutoLiquid_GenScript_Single_Handling.EntityJson
{
    /// <summary>
    /// 全部偏移值
    /// </summary>
    [Serializable()]
    public class Offsets
    {
        // 其他机器相对标准原点偏移量
        public List<Device> Devices = new List<Device>();
    }
}
