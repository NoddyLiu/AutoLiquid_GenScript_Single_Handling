using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using AutoLiquid_GenScript_Single_Handling.EntityCommon;
using AutoLiquid_GenScript_Single_Handling.EntityJson;
using AutoLiquid_Library.Enum;
using Path = System.Windows.Shapes.Path;

namespace AutoLiquid_GenScript_Single_Handling.Utils
{
    /// <summary>
    /// 耗材帮助类
    /// </summary>
    public class ConsumableHelper
    {
        /// <summary>
        /// 根据组名找出参数组
        /// </summary>
        /// <param name="headIndex">移液头Index</param>
        /// <param name="consumableTypeName">耗材类型</param>
        /// <param name="isTipBox">是否枪头盒</param>
        /// <returns></returns>
        public static Consumable GetConsumableType(int headIndex, string consumableTypeName, bool isTipBox)
        {
            var cGroup = ParamsHelper.CommonSettingList[headIndex].Consumables.FirstOrDefault(x => x.GroupName.Equals(consumableTypeName) && x.IsTipBox == isTipBox);
            return cGroup;
        }

        /// <summary>
        /// 获取孔Index（index规则：A1~H1 -> 0~8）
        /// </summary>
        /// <param name="headIndex">移液头Index</param>
        /// <param name="consumableType">耗材类型</param>
        /// <param name="holeIndexStr">孔字符，如A1、B3等</param>
        /// <returns></returns>
        public static HoleIndex GetHoleIndex(int headIndex, Consumable consumableType, string holeIndexStr)
        {
            var count = ParamsHelper.CommonSettingList[headIndex].A1Pos == EA1Pos.LeftTop ? consumableType.RowCount : consumableType.ColCount;
            var posLetter = holeIndexStr.Substring(0, 1);
            var posNum = holeIndexStr.Substring(1);
            var result = (Convert.ToInt32(posNum) - 1) * count + (posLetter.ToCharArray()[0] - 65);
            return new HoleIndex { OriIndex = result };
        }

        /// <summary>
        ///  根据index返回孔位置（例如返回A1、B2等）
        /// </summary>
        /// <param name="holeIndex"></param>
        /// <param name="rowCount"></param>
        /// <param name="colCount"></param>
        /// <param name="a1Pos"></param>
        /// <param name="indexHorizontal">index是否水平递增（默认垂直递增）</param>
        /// <returns></returns>
        public static string GetHolePosStr(int holeIndex, int rowCount, int colCount, EA1Pos a1Pos, bool indexHorizontal = false)
        {
            var result = "";
            if (a1Pos == EA1Pos.LeftTop)
            {
                if (!indexHorizontal)
                {
                    var rowIndex = holeIndex % rowCount;
                    var colIndex = holeIndex / rowCount;
                    result = ViewUtils.PosLetterList[rowIndex] + ViewUtils.PosNumList[colIndex];
                }
                else
                {
                    var colIndex = holeIndex % colCount;
                    var rowIndex = holeIndex / colCount;
                    result = ViewUtils.PosLetterList[rowIndex] + ViewUtils.PosNumList[colIndex];
                }
            }
            else
            {
                if (!indexHorizontal)
                {
                    var rowIndex = holeIndex / colCount;
                    var colIndex = holeIndex % colCount;
                    result = ViewUtils.PosLetterList[colIndex] + ViewUtils.PosNumList[rowIndex];
                }
                else
                {
                    var colIndex = holeIndex / colCount;
                    var rowIndex = holeIndex % colCount;
                    result = ViewUtils.PosLetterList[colIndex] + ViewUtils.PosNumList[rowIndex];
                }
            }

            return result;
        }



        /// <summary>
        /// 是否为384枪头盒（如果枪头盒x和y步长为4.5mm，且枪头盒总孔数为384，则认为是384枪头盒）
        /// </summary>
        /// <param name="xStep"></param>
        /// <param name="yStep"></param>
        /// <param name="rowCount"></param>
        /// <param name="colCount"></param>
        /// <returns></returns>
        public static bool Is384TipBox(decimal xStep, decimal yStep, int rowCount, int colCount)
        {
            return xStep == 4.5m && yStep == 4.5m && rowCount * colCount == 384;
        }
    }
}
