using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace AutoLiquid_GenScript_Single_Handling.EntityJson
{
    /// <summary>
    /// 轴校准
    /// </summary>
    [Serializable()]
    public class Calibration
    {
        // 是否启用
        public bool Available = false;

        // 需要校准的对应体积
        public decimal[] PVolArray = { 1000.0m, 500.0m, 300.0m, 100.0m, 10.0m };
        // 校准体积对应的补偿值
        public decimal[] PCompensationArray = { 0.0m, 0.0m, 0.0m, 0.0m, 0.0m };
    }
}
