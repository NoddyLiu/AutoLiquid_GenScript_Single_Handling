namespace AutoLiquid_Library.Enum
{
    /// <summary>
    /// 移液头状态
    /// </summary>
    public enum EHeadStatus
    {
        TipTook, // 已经取枪头
        Absorbed, // 已经吸液
        Jetted, // 已经喷液
        TipReleased // 已经退枪头
    }
}