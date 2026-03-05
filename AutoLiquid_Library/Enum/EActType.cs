using System.ComponentModel;

namespace AutoLiquid_Library.Enum
{
    /// <summary>
    /// 运动类型
    /// </summary>
    public enum EActType
    {
        // 绝对运动
        [Description("A")]
        A,
        // 相对运动
        [Description("S")]
        S,
        // 初始化
        [Description("I")]
        I,
        // 速度设置
        [Description("F")]
        F,
        // 中途停止
        [Description("W")]
        W,
        // 往复运动
        [Description("V")]
        V,
    }
}