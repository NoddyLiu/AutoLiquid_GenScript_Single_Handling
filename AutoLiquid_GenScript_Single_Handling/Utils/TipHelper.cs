using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using AutoLiquid_GenScript_Single_Handling.EntityCommon;
using AutoLiquid_GenScript_Single_Handling.EntityJson;
using AutoLiquid_GenScript_Single_Handling.UserControls;
using AutoLiquid_Library.Enum;
using AutoLiquid_Library.Utils;
using Newtonsoft.Json;

namespace AutoLiquid_GenScript_Single_Handling.Utils
{
    /// <summary>
    /// 枪头工具类
    /// </summary>
    public class TipHelper
    {
        /// <summary>
        /// 获取枪头数目每行数据
        /// </summary>
        /// <param name="matrix"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public static bool[] GetTipChannel2DArrayRow(bool[,] matrix, int rowIndex)
        {
            return Enumerable.Range(0, matrix.GetLength(1))
                .Select(x => matrix[rowIndex, x])
                .ToArray();
        }

        /// <summary>
        /// 获取枪头数目每列数据
        /// </summary>
        /// <param name="matrix"></param>
        /// <param name="colIndex"></param>
        /// <returns></returns>
        public static bool[] GetTipChannel2DArrayColumn(bool[,] matrix, int colIndex)
        {
            return Enumerable.Range(0, matrix.GetLength(0))
                .Select(x => matrix[x, colIndex])
                .ToArray();
        }

        /// <summary>
        /// 枪头盒是否灵活取枪头
        /// </summary>
        /// <param name="headUsedIndex"></param>
        /// <param name="seqList"></param>
        /// <param name="tipTemplateConsumable"></param>
        /// <returns></returns>
        public static bool IsTipBoxFlexible(int headUsedIndex, List<Seq> seqList, Consumable tipTemplateConsumable)
        {
            return seqList.Exists(p => IsTakeTipFlexible(headUsedIndex, p.TipChannel, tipTemplateConsumable));
        }

        /// <summary>
        /// 是否灵活取枪头
        /// </summary>
        /// <param name="headUsedIndex"></param>
        /// <param name="tipChannel2DArray"></param>
        /// <param name="tipTemplateConsumable"></param>
        /// <returns></returns>
        public static bool IsTakeTipFlexible(int headUsedIndex, int[,] tipChannel2DArray, Consumable tipTemplateConsumable)
        {
            // 移液头通道数
            var headChannelRow = ParamsHelper.HeadList[headUsedIndex].ChannelRow;
            var headChannelCol = ParamsHelper.HeadList[headUsedIndex].ChannelCol;

            // 所需枪头行、列数
            var tipChannelRow = tipChannel2DArray.GetLength(0);
            var tipChannelCol = tipChannel2DArray.GetLength(1);

            /**
             * 判断逻辑：满足以下条件之一，则认为是灵活取枪头，取枪头位置需要偏移
             * 1、移液头通道数和所需枪头通道数是否一致
             * 2、非单通道移液头所需枪头是否与枪头盒耗材行列一致（即所需枪头行列均与枪头盒耗材行列不一样）
             */
            // return headChannelRow != tipChannelRow || headChannelCol != tipChannelCol
            //                                        || (headChannelRow * headChannelCol != 1 && tipChannelRow != tipTemplateConsumable.RowCount && tipChannelCol != tipTemplateConsumable.ColCount);

            return headChannelRow != tipChannelRow || headChannelCol != tipChannelCol;
        }

        /// <summary>
        /// 计算移液头所需枪头数目
        /// </summary>
        /// <param name="seqList">移液序列</param>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="tipTemplateIndex">非null：指定枪头盒盘位</param>
        /// <param name="tipCount2Need">所需枪头数</param>
        /// <param name="replaceTipOverRange">超过量程移液中途是否换枪头</param>
        public static void CalculateTipNeedCountByHead(List<Seq> seqList, int headUsedIndex, int? tipTemplateIndex, ref int tipCount2Need, bool replaceTipOverRange)
        {
            // 移液头最大量程
            var headLiquidRangeMax = ConstantsUtils.LiquidRangeMaxDic[ParamsHelper.HeadList[headUsedIndex].HeadLiquidRange];
            // 移液头实际量程
            var headLiquidRangeReal = Convert.ToDecimal((int)ParamsHelper.HeadList[headUsedIndex].HeadLiquidRange);
            // 吸液前吸空气量
            var absorbAirBefore = headLiquidRangeReal * ParamsHelper.CommonSettingList[headUsedIndex].AbsorbAirBeforePercent * 0.01m;
            // 吸液后吸空气量
            var absorbAirAfter = headLiquidRangeReal * ParamsHelper.CommonSettingList[headUsedIndex].AbsorbAirAfterPercent * 0.01m;
            // // 一吸多喷吸液后多吸体积
            // var absorbLiquidMoreOne2More = headLiquidRangeReal * ParamsHelper.CommonSettingList[headUsedIndex].AbsorbLiquidMoreOne2MorePercent * 0.01m;
            // 吸液后多吸体积
            var absorbLiquidMore = headLiquidRangeReal * ParamsHelper.CommonSettingList[headUsedIndex].AbsorbLiquidMorePercent * 0.01m;
            // 额外多吸体积
            var additionalVolume = absorbAirBefore + absorbAirAfter;

            /**
             * 逻辑：
             * ①正常情况下（超过量程移液中途不换枪头）
             * ②非正常情况下（超过量程移液中途换枪头）
             */
            // ①正常情况下（超过量程移液中途不换枪头）
            var seqsNormal = tipTemplateIndex == null ? seqList.Where(p => p.IsTakeTip && p.HeadUsedIndex == headUsedIndex && !p.IsComment) : seqList.Where(p => p.IsTakeTip && p.HeadUsedIndex == headUsedIndex && !p.IsComment && p.TipTemplateIndex == tipTemplateIndex);
            foreach (var seq in seqsNormal)
            {
                var tipChannel = seq.TipChannel.GetLength(0) * seq.TipChannel.GetLength(1);
                tipCount2Need += tipChannel;
            }

            // ②非正常情况下（超过量程移液中途换枪头）
            if (replaceTipOverRange)
            {
                /**
                * 一吸一喷
                */
                // 一吸一喷中，计算单条分液过程中使用枪头数目
                var seqListWhereOne2One = tipTemplateIndex == null ? seqList.Where(p => p.TargetHoleIndexList.Count == 1 && p.HeadUsedIndex == headUsedIndex && !p.IsComment) : seqList.Where(p => p.TargetHoleIndexList.Count == 1 && p.HeadUsedIndex == headUsedIndex && !p.IsComment && p.TipTemplateIndex == tipTemplateIndex);
                for (var i = 0; i < seqListWhereOne2One.Count(); i++)
                {
                    var seq = seqListWhereOne2One.ElementAt(i);

                    var absorbCountInThisSeqOne2One = 1; // 分液次数
                    var liquidVolumeAbsorbEach = seq.VolumeEachList[0].Calibration;   // 每次吸液量
                    var liquidVolumeTotal = liquidVolumeAbsorbEach; // 一共吸取的体积
                    // 额外多吸体积
                    additionalVolume += absorbLiquidMore;
                    // 注意：需要加上空气量再比较，（吸液前后空气量 + 吸液量）是否大于量程最大值
                    while (liquidVolumeAbsorbEach + additionalVolume > headLiquidRangeMax)
                    {
                        absorbCountInThisSeqOne2One++;
                        liquidVolumeAbsorbEach = liquidVolumeTotal / absorbCountInThisSeqOne2One;
                    }
                    // 叠加总枪头数目
                    var tipChannel = seq.TipChannel.GetLength(0) * seq.TipChannel.GetLength(1);
                    tipCount2Need += (absorbCountInThisSeqOne2One - 1) * tipChannel;
                }
                /**
                 * 一吸多喷
                 */
                // 一吸多喷中，计算单条分液过程中使用枪头数目
                var seqListWhereOne2More = tipTemplateIndex == null ? seqList.Where(p => p.TargetHoleIndexList.Count > 1 && p.HeadUsedIndex == headUsedIndex && !p.IsComment) : seqList.Where(p => p.TargetHoleIndexList.Count > 1 && p.HeadUsedIndex == headUsedIndex && !p.IsComment && p.TipTemplateIndex == tipTemplateIndex);
                for (var i = 0; i < seqListWhereOne2More.Count(); i++)
                {
                    var seq = seqListWhereOne2More.ElementAt(i);

                    var liquidVolumeEach = 0m;   // 每次吸液量
                    var absorbCountInThisSeqOne2More = 0; // 本序列分液次数
                    var lastHoleOverRange = false; // 上一孔是否超出量程（用于判断本孔与上一孔比较，如果上一孔超出量程，则本孔分液次数需要+1）
                    for (var index = 0; index < seq.TargetHoleIndexList.Count; index++)
                    {
                        // 当前孔吸液体积
                        var volumeCurrentHole = seq.VolumeEachList.Count == 1
                            ? seq.VolumeEachList[0].Calibration
                            : seq.VolumeEachList[index].Calibration;

                        /**
                         * 单孔超过量程，直接平均分液（即在本孔容量基础上直接平均分液）；否则叠加分液
                         */
                        // 如果当前孔吸液体积 > 量程，直接计算该孔需要移液次数；否则叠加计算移液次数
                        if (volumeCurrentHole + additionalVolume > headLiquidRangeMax)
                        {
                            var absorbCountInThisHoleOne2More = 1; // 本孔分液次数
                            liquidVolumeEach = volumeCurrentHole;
                            while (liquidVolumeEach + additionalVolume > headLiquidRangeMax)
                            {
                                absorbCountInThisHoleOne2More++;
                                liquidVolumeEach = volumeCurrentHole / absorbCountInThisHoleOne2More;
                            }
                            // 把当前容量计数置零
                            liquidVolumeEach = 0;
                            absorbCountInThisSeqOne2More += absorbCountInThisHoleOne2More;
                            lastHoleOverRange = true;
                        }
                        else
                        {
                            liquidVolumeEach += volumeCurrentHole;

                            // 如果上一孔超过量程 或者 本孔为首孔，递增一次分液次数
                            if (lastHoleOverRange || index == 0)
                                absorbCountInThisSeqOne2More++;

                            // 注意：需要加上空气量再比较，（吸液前后空气量 + 吸液量）是否大于量程最大值
                            if (liquidVolumeEach + additionalVolume > headLiquidRangeMax)
                            {
                                absorbCountInThisSeqOne2More++;
                                liquidVolumeEach = 0;
                                index -= 1; // 从上一个靶孔开始检索
                            }

                            lastHoleOverRange = false;
                        }
                    }

                    // 叠加总枪头数目
                    var tipChannel = seq.TipChannel.GetLength(0) * seq.TipChannel.GetLength(1);
                    tipCount2Need += (absorbCountInThisSeqOne2More - 1) * tipChannel;
                }
            }
        }

        /// <summary>
        /// 计算枪头盒剩余枪头数目
        /// </summary>
        /// <param name="controlTemplateTip">枪头盒盘位</param>
        /// <returns></returns>
        public static int CalculateTipRemainCountByHead(ControlTemplateTip controlTemplateTip)
        {
            var result = 0;

            var tipChannelRow = controlTemplateTip.TipBoxUsedStatus2DArray.GetLength(0);
            var tipChannelCol = controlTemplateTip.TipBoxUsedStatus2DArray.GetLength(1);
            for (var row = 0; row < tipChannelRow; row++)
            {
                for (var col = 0; col < tipChannelCol; col++)
                {
                    if (!controlTemplateTip.TipBoxUsedStatus2DArray[row, col])
                        result++;
                }
            }

            return result;
        }

        /// <summary>
        /// 计算本次取枪头位置，并返回下一次取枪头位置
        /// </summary>
        /// <param name="headUsedIndex"></param>
        /// <param name="tipTemplateConsumable">枪头盒耗材</param>
        /// <param name="tipChannel2DArray">枪头使用数量二维数组</param>
        /// <param name="allTipTemplateList">全部枪头盘位</param>
        /// <param name="tipTemplateUsedIndex">用到的枪头盘位Index</param>
        /// <param name="tipUsedIndexThisTimeList">本次所用的枪头行、列index列表</param>
        /// <param name="tipTakeStartIndex">取枪头开始行、列Index</param>
        /// <param name="isTipCountEnough">是否有足够的枪头数</param>
        /// <returns>下一次取枪头位置</returns>
        public static int CalcTipPos(int headUsedIndex, Consumable tipTemplateConsumable, int[,] tipChannel2DArray, List<ControlTemplateTip> allTipTemplateList, ref int tipTemplateUsedIndex, ref List<RowCol> tipUsedIndexThisTimeList, ref HoleIndex tipTakeStartIndex, ref bool isTipCountEnough)
        {
            // 是否终止顶层for循环
            var isBreakParentLoop = false;

            // y轴移动系数（正方向：头移动；反方向：盘移动）
            var yDirectionFactor = ParamsHelper.HeadList[headUsedIndex].YMoveWithHead ? 1 : -1;

            // 所需枪头行、列数
            var tipChannelRow = tipChannel2DArray.GetLength(0);
            var tipChannelCol = tipChannel2DArray.GetLength(1);

            // 移液头通道数
            var headChannelRow = ParamsHelper.HeadList[headUsedIndex].ChannelRow;
            var headChannelCol = ParamsHelper.HeadList[headUsedIndex].ChannelCol;
            // A1摆放位置
            var a1Pos = ParamsHelper.CommonSettingList[headUsedIndex].A1Pos;
            // 逐列取枪头
            var takeTipEachCol = ParamsHelper.CommonSettingList[headUsedIndex].TakeTipEachCol;
            // 取枪头方向：从左往右
            var takeTipLeft2Right = ParamsHelper.CommonSettingList[headUsedIndex].TakeTipLeft2Right || headChannelRow * headChannelCol == 1;

            // 是否灵活取枪头
            bool isTakeTipFlexible = IsTakeTipFlexible(headUsedIndex, tipChannel2DArray, tipTemplateConsumable);
            // 移液头通道间距是否标准9.0mm
            var channelStepStandard = ParamsHelper.HeadList[headUsedIndex].ChannelStep == 9.0m;
            // 枪头盒孔间距
            var tipBoxHoleStep = tipTemplateConsumable.HoleStep;
            // 枪头盒行列数
            var tipBoxRowCount = tipTemplateConsumable.RowCount;
            var tipBoxColCount = tipTemplateConsumable.ColCount;

            // 下一次取枪头位置
            var nextTipIndex = 0;

            // 遍历每个枪头盒
            for (var i = 0; i < allTipTemplateList.Count; i++)
            {
                if (isBreakParentLoop)
                    break;

                var tipTemplate = allTipTemplateList[i];
                var tipBoxUsedStatusRow = tipTemplate.TipBoxUsedStatus2DArray.GetLength(0);
                var tipBoxUsedStatusCol = tipTemplate.TipBoxUsedStatus2DArray.GetLength(1);
                var totalHoles = tipBoxUsedStatusRow * tipBoxUsedStatusCol;

                // A1左上
                if (a1Pos == EA1Pos.LeftTop)
                {
                    // 逐列取
                    if (takeTipEachCol)
                    {
                        // 找出每列可用枪头数目，判断是否足够，不足的话，到下一列再找
                        for (var tipBoxCol = takeTipLeft2Right ? 0 : tipBoxUsedStatusCol - 1; takeTipLeft2Right ? tipBoxCol < tipBoxUsedStatusCol : tipBoxCol >= 0; tipBoxCol = takeTipLeft2Right ? ++tipBoxCol : --tipBoxCol)
                        {
                            // 该列剩余枪头数组
                            var tipRemainArrayInThisCol = GetTipChannel2DArrayColumn(tipTemplate.TipBoxUsedStatus2DArray, tipBoxCol);
                            // 该列剩余的枪头数目
                            var tipRemainCountInThisCol = tipRemainArrayInThisCol.Count(p => !p);
                            // 是否有足够的枪头
                            if (tipRemainCountInThisCol < tipChannelRow)
                                continue;

                            isTipCountEnough = true;

                            /**
                            * 符合条件的枪头
                            */
                            // 开始取枪头的Row Index
                            // TODO 新修改：2024-03-07
                            var tipBoxUnusedIndexList = GetTipBoxUnusedIndexList(headUsedIndex, tipChannel2DArray, tipTemplateConsumable, tipRemainArrayInThisCol, tipChannelRow, takeTipLeft2Right);
                            var tipStartRowIndex = tipBoxUnusedIndexList.First();

                            // 本次所用的枪头index（如果需要多列枪头，后续列也一并使用）
                            for (var tipCol = tipBoxCol; takeTipLeft2Right ? tipCol < tipBoxCol + tipChannelCol : tipCol > tipBoxCol - tipChannelCol; tipCol = takeTipLeft2Right ? ++tipCol : --tipCol)
                            {
                                // TODO 新修改：2024-03-07
                                foreach (var tipRow in tipBoxUnusedIndexList)
                                {
                                    var currentRowCol = new RowCol { Row = tipRow, Col = tipCol };
                                    tipUsedIndexThisTimeList.Add(currentRowCol);
                                    var isLastRowIndex = tipRow == tipBoxUsedStatusRow - 1;
                                    if (takeTipLeft2Right)
                                        nextTipIndex = currentRowCol.Col * tipBoxUsedStatusRow + currentRowCol.Row + 1;
                                    else
                                        nextTipIndex = isLastRowIndex ? (currentRowCol.Col - 1) * tipBoxUsedStatusRow : currentRowCol.Col * tipBoxUsedStatusRow + currentRowCol.Row + 1;
                                }
                            }

                            tipTemplateUsedIndex = tipTemplate.TipBoxTemplateIndex;

                            var oriIndex = takeTipLeft2Right ? tipUsedIndexThisTimeList.First().Col * tipBoxUsedStatusRow : tipUsedIndexThisTimeList.Last().Col * tipBoxUsedStatusRow;
                            // 取枪头通道数与移液头通道数一致，取枪头位置不需要偏移
                            if (!isTakeTipFlexible)
                                tipTakeStartIndex = new HoleIndex { OriIndex = oriIndex + tipStartRowIndex };
                            else
                            {
                                /**
                                * 移液头偏移逻辑：
                                * ①移液头多列，整列取：X轴偏移
                                * ②移液头多列，灵活取：X、Y轴偏移
                                * ③移液头单列，Y轴偏移
                                */
                                var xHoleOffset = takeTipLeft2Right ? tipChannelCol - headChannelCol : 0;
                                var yHoleOffset = (tipStartRowIndex - (tipBoxUsedStatusRow - tipChannelRow)) * yDirectionFactor;
                                if (headChannelCol > 1)
                                {
                                    if (tipChannelRow == headChannelRow)
                                        tipTakeStartIndex = new HoleIndex { OriIndex = oriIndex, XHoleOffset = xHoleOffset };
                                    else
                                        tipTakeStartIndex = new HoleIndex { OriIndex = oriIndex, XHoleOffset = xHoleOffset, YHoleOffset = yHoleOffset };
                                }
                                else
                                    tipTakeStartIndex = new HoleIndex { OriIndex = oriIndex, YHoleOffset = yHoleOffset };
                            }

                            isBreakParentLoop = true;
                            break;
                        }
                    }
                    // 逐行取
                    else
                    {
                        // 找出每行可用枪头数目，判断是否足够，不足的话，到下一行再找
                        for (var tipBoxRow = 0; tipBoxRow < tipBoxUsedStatusRow; tipBoxRow++)
                        {
                            // 该行剩余枪头数组
                            var tipRemainArrayInThisRow = GetTipChannel2DArrayRow(tipTemplate.TipBoxUsedStatus2DArray, tipBoxRow);
                            // 该行剩余的枪头数目
                            var tipRemainCountInThisRow = tipRemainArrayInThisRow.Count(p => !p);
                            // 是否有足够的枪头
                            if (tipRemainCountInThisRow < tipChannelCol)
                                continue;

                            isTipCountEnough = true;

                            /**
                            * 符合条件的枪头
                            */
                            // 开始取枪头的Col Index
                            // TODO 新修改：2024-03-07
                            var tipBoxUnusedIndexList = GetTipBoxUnusedIndexList(headUsedIndex, tipChannel2DArray, tipTemplateConsumable, tipRemainArrayInThisRow, tipChannelCol, takeTipLeft2Right);
                            var tipStartColIndex = tipBoxUnusedIndexList.First();

                            // 本次所用的枪头index（如果需要多行枪头，后续行也一并使用）
                            // for (var tipRow = tipBoxRow; tipRow < tipBoxRow + tipChannelRow; tipRow++)
                            // {
                            //     // TODO 新修改：2024-03-07
                            //     foreach (var tipCol in tipBoxUnusedIndexList)
                            //     {
                            //         var currentRowCol = new RowCol { Row = tipRow, Col = tipCol };
                            //         tipUsedIndexThisTimeList.Add(currentRowCol);
                            //         var isLastColIndex = takeTipLeft2Right ? tipCol == tipBoxUsedStatusCol - 1 : tipCol == 0;
                            //         if (takeTipLeft2Right)
                            //             nextTipIndex = isLastColIndex ? (currentRowCol.Row + 1 == tipBoxUsedStatusRow ? -1 : currentRowCol.Row + 1) : (currentRowCol.Col + 1) * tipBoxUsedStatusRow + currentRowCol.Row;
                            //         else
                            //             nextTipIndex = isLastColIndex ? (tipBoxUsedStatusCol - 1) * tipBoxUsedStatusRow + currentRowCol.Row + 1 : (currentRowCol.Col - 1) * tipBoxUsedStatusRow + currentRowCol.Row;
                            //     }
                            // }
                            // 本次所用的枪头index（如果需要多行枪头，后续行也一并使用）
                            for (var rowSeq = 0; rowSeq < tipChannelRow; rowSeq++)
                            {
                                // 如果枪头盒为384枪头盒，且移液头通道间距为标准9.0mm，则多行或者多列取枪头时，需要间隔取
                                var holeMismatch = ConsumableHelper.Is384TipBox(tipBoxHoleStep.X, tipBoxHoleStep.Y, tipBoxRowCount, tipBoxColCount) && channelStepStandard;
                                // 间隔取步长
                                var mismatchStep = holeMismatch ? Convert.ToInt16(TipHelper.MultiChannelStepAndConsumableStepRelation(headUsedIndex, tipChannel2DArray, tipTemplateConsumable)) : 1;
                                var tipRow = tipBoxRow + rowSeq * mismatchStep;
                                foreach (var tipCol in tipBoxUnusedIndexList)
                                {
                                    var currentRowCol = new RowCol { Row = tipRow, Col = tipCol };
                                    tipUsedIndexThisTimeList.Add(currentRowCol);
                                    var isLastColIndex = takeTipLeft2Right ? tipCol == tipBoxUsedStatusCol - 1 : tipCol == 0;
                                    if (takeTipLeft2Right)
                                        // nextTipIndex = isLastColIndex ? (currentRowCol.Row + 1 == tipBoxUsedStatusRow ? -1 : currentRowCol.Row + 1) : (currentRowCol.Col + 1) * tipBoxUsedStatusRow + currentRowCol.Row;
                                        // 到达行末时，跳到下一行首列；整盒用完时返回 -1
                                        nextTipIndex = isLastColIndex
                                            ? (currentRowCol.Row + 1 == tipBoxUsedStatusRow
                                                ? -1
                                                : (currentRowCol.Row + 1) * tipBoxUsedStatusCol)
                                            : currentRowCol.Row * tipBoxUsedStatusCol + currentRowCol.Col + 1;

                                    else
                                        // nextTipIndex = isLastColIndex ? (tipBoxUsedStatusCol - 1) * tipBoxUsedStatusRow + currentRowCol.Row + 1 : (currentRowCol.Col - 1) * tipBoxUsedStatusRow + currentRowCol.Row;
                                        nextTipIndex = isLastColIndex ? (currentRowCol.Row + 1 == tipBoxUsedStatusRow ? -1 : currentRowCol.Row + 1) : (currentRowCol.Col - 1) * tipBoxUsedStatusRow + currentRowCol.Row;
                                }
                            }

                            tipTemplateUsedIndex = tipTemplate.TipBoxTemplateIndex;

                            // 取枪头通道数与移液头通道数一致，取枪头位置不需要偏移
                            if (!isTakeTipFlexible)
                                tipTakeStartIndex = new HoleIndex { OriIndex = takeTipLeft2Right ? tipUsedIndexThisTimeList.First().Col * tipBoxUsedStatusRow + tipBoxRow : tipUsedIndexThisTimeList.Last().Col * tipBoxUsedStatusRow + tipBoxRow };
                            else
                            {
                                /**
                                 * 移液头偏移逻辑：
                                 * ①移液头多行，整行取：Y轴偏移
                                 * ②移液头多行，灵活取：X、Y轴偏移
                                 * ③移液头单行，X轴偏移
                                 */
                                var xHoleOffset = takeTipLeft2Right ? tipStartColIndex - (tipBoxUsedStatusCol - tipChannelCol) : tipStartColIndex - tipChannelCol + 1;
                                var yHoleOffset = (tipChannelRow - headChannelRow) * yDirectionFactor;
                                if (headChannelRow > 1)
                                {
                                    if (tipChannelCol == headChannelCol)
                                        tipTakeStartIndex = new HoleIndex { OriIndex = tipBoxRow, YHoleOffset = yHoleOffset };
                                    else
                                        tipTakeStartIndex = new HoleIndex { OriIndex = tipBoxRow, XHoleOffset = xHoleOffset, YHoleOffset = yHoleOffset };
                                }
                                else
                                    tipTakeStartIndex = new HoleIndex { OriIndex = tipBoxRow, XHoleOffset = xHoleOffset };
                            }
                            isBreakParentLoop = true;
                            break;
                        }
                    }
                }
                // A1左下
                else
                {
                    // 逐列取
                    if (takeTipEachCol)
                    {
                        // 找出每列可用枪头数目，判断是否足够，不足的话，到下一列再找
                        for (var tipBoxCol = takeTipLeft2Right ? 0 : tipBoxUsedStatusCol - 1; takeTipLeft2Right ? tipBoxCol < tipBoxUsedStatusCol : tipBoxCol >= 0; tipBoxCol = takeTipLeft2Right ? ++tipBoxCol : --tipBoxCol)
                        {
                            // 该列剩余枪头数组
                            var tipRemainArrayInThisCol = GetTipChannel2DArrayColumn(tipTemplate.TipBoxUsedStatus2DArray, tipBoxCol);
                            // 该列剩余的枪头数目
                            var tipRemainCountInThisCol = tipRemainArrayInThisCol.Count(p => !p);
                            // 是否有足够的枪头
                            if (tipRemainCountInThisCol < tipChannelRow)
                                continue;

                            isTipCountEnough = true;

                            /**
                            * 符合条件的枪头
                            */
                            // 开始取枪头的Row Index
                            // TODO 新修改：2024-03-07
                            var tipBoxUnusedIndexList = GetTipBoxUnusedIndexList(headUsedIndex, tipChannel2DArray, tipTemplateConsumable, tipRemainArrayInThisCol, tipChannelRow, takeTipLeft2Right);
                            var tipStartRowIndex = tipBoxUnusedIndexList.First();

                            // 本次所用的枪头index（如果需要多列枪头，后续列也一并使用）
                            for (var tipCol = tipBoxCol; takeTipLeft2Right ? tipCol < tipBoxCol + tipChannelCol : tipCol > tipBoxCol - tipChannelCol; tipCol = takeTipLeft2Right ? ++tipCol : --tipCol)
                            {
                                // TODO 新修改：2024-03-07
                                foreach (var tipRow in tipBoxUnusedIndexList)
                                {
                                    var currentRowCol = new RowCol { Row = tipRow, Col = tipCol };
                                    tipUsedIndexThisTimeList.Add(currentRowCol);

                                    var isLastRowIndex = tipRow == tipBoxUsedStatusRow - 1;
                                    if (takeTipLeft2Right)
                                        nextTipIndex = isLastRowIndex ? (currentRowCol.Col + 1 == tipBoxUsedStatusCol ? -1 : currentRowCol.Col + 1) : (currentRowCol.Row + 1) * tipBoxUsedStatusCol + currentRowCol.Col;
                                    else
                                        nextTipIndex = isLastRowIndex ? currentRowCol.Col - 1 : (currentRowCol.Row + 1) * tipBoxUsedStatusCol + currentRowCol.Col;
                                }
                            }

                            tipTemplateUsedIndex = tipTemplate.TipBoxTemplateIndex;

                            // 取枪头通道数与移液头通道数一致，取枪头位置不需要偏移
                            if (!isTakeTipFlexible)
                                tipTakeStartIndex = new HoleIndex { OriIndex = takeTipLeft2Right ? tipUsedIndexThisTimeList.First().Row * tipBoxUsedStatusCol + tipBoxCol : tipBoxCol + 1 - tipChannelCol };
                            else
                            {
                                /**
                                * 移液头偏移逻辑：
                                * ①移液头多列，整列取：X轴偏移
                                * ②移液头多列，灵活取：X、Y轴偏移
                                * ③移液头单列，Y轴偏移
                                */
                                var xHoleOffset = takeTipLeft2Right ? tipChannelCol - headChannelCol : tipUsedIndexThisTimeList.Last().Col;
                                var yHoleOffset = (tipBoxUsedStatusRow - tipChannelRow - tipStartRowIndex) * yDirectionFactor;
                                if (headChannelCol > 1)
                                {
                                    if (tipChannelRow == headChannelRow)
                                        tipTakeStartIndex = new HoleIndex { OriIndex = takeTipLeft2Right ? tipBoxCol : 0, XHoleOffset = xHoleOffset };
                                    else
                                        tipTakeStartIndex = new HoleIndex { OriIndex = takeTipLeft2Right ? tipBoxCol : 0, XHoleOffset = xHoleOffset, YHoleOffset = yHoleOffset };
                                }
                                else
                                    tipTakeStartIndex = new HoleIndex { OriIndex = tipBoxCol, YHoleOffset = yHoleOffset };
                            }

                            isBreakParentLoop = true;
                            break;
                        }
                    }
                    // 逐行取
                    else
                    {
                        // 找出每行可用枪头数目，判断是否足够，不足的话，到下一行再找
                        for (var tipBoxRow = 0; tipBoxRow < tipBoxUsedStatusRow; tipBoxRow++)
                        {
                            // 该行剩余枪头数组
                            var tipRemainArrayInThisRow = GetTipChannel2DArrayRow(tipTemplate.TipBoxUsedStatus2DArray, tipBoxRow);
                            // 该行剩余的枪头数目
                            var tipRemainCountInThisRow = tipRemainArrayInThisRow.Count(p => !p);
                            // 是否有足够的枪头
                            if (tipRemainCountInThisRow < tipChannelCol)
                                continue;

                            isTipCountEnough = true;

                            /**
                            * 符合条件的枪头
                            */
                            // 开始取枪头的Col Index
                            // TODO 新修改：2024-03-07
                            var tipBoxUnusedIndexList = GetTipBoxUnusedIndexList(headUsedIndex, tipChannel2DArray, tipTemplateConsumable, tipRemainArrayInThisRow, tipChannelCol, takeTipLeft2Right);
                            var tipStartColIndex = tipBoxUnusedIndexList.First();

                            // 本次所用的枪头index（如果需要多行枪头，后续行也一并使用）
                            // for (var tipRow = tipBoxRow; tipRow < tipBoxRow + tipChannelRow; tipRow++)
                            // {
                            //     // TODO 新修改：2024-03-07
                            //     foreach (var tipCol in tipBoxUnusedIndexList)
                            //     {
                            //         var currentRowCol = new RowCol { Row = tipRow, Col = tipCol };
                            //         tipUsedIndexThisTimeList.Add(currentRowCol);
                            //         
                            //         var isLastColIndex = takeTipLeft2Right ? tipCol == tipBoxUsedStatusCol - 1 : tipCol == 0;
                            //         if (takeTipLeft2Right)
                            //             nextTipIndex = currentRowCol.Row * tipBoxUsedStatusCol + currentRowCol.Col + 1;
                            //         else
                            //             nextTipIndex = isLastColIndex ? (currentRowCol.Row + 2) * tipBoxUsedStatusCol - 1 : currentRowCol.Row * tipBoxUsedStatusCol + currentRowCol.Col - 1;
                            //     }
                            // }
                            // 本次所用的枪头index（如果需要多行枪头，后续行也一并使用）
                            for (var rowSeq = 0; rowSeq < tipChannelRow; rowSeq++)
                            {
                                // 如果枪头盒为384枪头盒，且移液头通道间距为标准9.0mm，则多行或者多列取枪头时，需要间隔取
                                var holeMismatch = ConsumableHelper.Is384TipBox(tipBoxHoleStep.X, tipBoxHoleStep.Y, tipBoxRowCount, tipBoxColCount) && channelStepStandard;
                                // 间隔取步长
                                var mismatchStep = holeMismatch ? Convert.ToInt16(TipHelper.MultiChannelStepAndConsumableStepRelation(headUsedIndex, tipChannel2DArray, tipTemplateConsumable)) : 1;
                                var tipRow = tipBoxRow + rowSeq * mismatchStep;
                                foreach (var tipCol in tipBoxUnusedIndexList)
                                {
                                    var currentRowCol = new RowCol { Row = tipRow, Col = tipCol };
                                    tipUsedIndexThisTimeList.Add(currentRowCol);

                                    var isLastColIndex = takeTipLeft2Right ? tipCol == tipBoxUsedStatusCol - 1 : tipCol == 0;
                                    if (takeTipLeft2Right)
                                        nextTipIndex = currentRowCol.Row * tipBoxUsedStatusCol + currentRowCol.Col + 1;
                                    else
                                        nextTipIndex = isLastColIndex ? (currentRowCol.Row + 1) * tipBoxUsedStatusCol : currentRowCol.Row * tipBoxUsedStatusCol + currentRowCol.Col - 1;
                                }
                            }

                            tipTemplateUsedIndex = tipTemplate.TipBoxTemplateIndex;

                            var oriIndex = tipUsedIndexThisTimeList.First().Row * tipBoxUsedStatusCol;
                            // 取枪头通道数与移液头通道数一致，取枪头位置不需要偏移
                            if (!isTakeTipFlexible)
                                tipTakeStartIndex = new HoleIndex { OriIndex = takeTipLeft2Right ? oriIndex + tipUsedIndexThisTimeList.First().Col : oriIndex + tipUsedIndexThisTimeList.Last().Col };
                            else
                            {
                                /**
                                 * 移液头偏移逻辑：
                                 * ①移液头多行，整行取：Y轴偏移
                                 * ②移液头多行，灵活取：X、Y轴偏移
                                 * ③移液头单行，X轴偏移
                                 */
                                var xHoleOffset = takeTipLeft2Right ? tipStartColIndex - (tipBoxUsedStatusCol - tipChannelCol) : tipStartColIndex - tipChannelCol + 1;
                                var yHoleOffset = (headChannelRow - tipChannelRow) * yDirectionFactor;
                                if (headChannelRow > 1)
                                {
                                    if (tipChannelCol == headChannelCol)
                                        tipTakeStartIndex = new HoleIndex { OriIndex = oriIndex, YHoleOffset = yHoleOffset };
                                    else
                                        tipTakeStartIndex = new HoleIndex { OriIndex = oriIndex, XHoleOffset = xHoleOffset, YHoleOffset = yHoleOffset };
                                }
                                else
                                    tipTakeStartIndex = new HoleIndex { OriIndex = oriIndex, XHoleOffset = xHoleOffset };
                            }

                            isBreakParentLoop = true;
                            break;
                        }
                    }
                }

                if (nextTipIndex >= totalHoles || nextTipIndex < 0)
                    nextTipIndex = -1;
            }

            return nextTipIndex;
        }

        /// <summary>
        /// 获取取枪头个数二维数组
        /// </summary>
        /// <param name="headUsedIndex"></param>
        /// <param name="tipChannel">取枪头个数</param>
        /// <returns></returns>
        public static int[,] GetTipChannel2DArray(int headUsedIndex, int tipChannel)
        {
            int[,] result = new int[1, 1];

            var channelRow = ParamsHelper.HeadList[headUsedIndex].ChannelRow;
            var channelCol = ParamsHelper.HeadList[headUsedIndex].ChannelCol;

            // 逐列取枪头
            var takeTipEachCol = ParamsHelper.CommonSettingList[headUsedIndex].TakeTipEachCol;

            // 如果取枪头个数大于总通道数，默认返回移液头通道数
            if (tipChannel > channelRow * channelCol)
            {
                return new int[channelRow, channelCol];
            }

            // 移液头是单行或者单列
            if (channelRow == 1 || channelCol == 1)
            {
                if (channelRow == 1)
                {
                    result = new int[1, tipChannel];
                }
                if (channelCol == 1)
                {
                    result = new int[tipChannel, 1];
                }
            }
            // 移液头是多行多列
            else
            {
                if (takeTipEachCol)
                {
                    // 判断是否整列取
                    var rowRemainder = tipChannel / channelRow;
                    var colArray = rowRemainder >= 1 ? rowRemainder : rowRemainder + 1;
                    result = new int[tipChannel <= channelRow ? tipChannel : channelRow, colArray];
                }
                else
                {
                    // 判断是否整行取
                    var colRemainder = tipChannel / channelCol;
                    var rowArray = colRemainder >= 1 ? colRemainder : colRemainder + 1;
                    result = new int[rowArray, tipChannel <= channelCol ? tipChannel : channelCol];
                }
            }

            return result;
        }

        /// <summary>
        /// 多通道移液头通道间距与耗材间距倍数关系
        /// </summary>
        /// <param name="headUsedIndex"></param>
        /// <param name="tipChannel2DArray">枪头使用数量二维数组</param>
        /// <param name="consumable">需要对比的耗材</param>
        /// <returns></returns>
        public static decimal MultiChannelStepAndConsumableStepRelation(int headUsedIndex, int[,] tipChannel2DArray, Consumable consumable)
        {
            var relation = 1.0m;

            var head = ParamsHelper.HeadList[headUsedIndex];

            // 如果为可变距，直接返回relation
            if (head.IsVariable)
                return relation;

            var channelRow = head.ChannelRow;
            var channelCol = head.ChannelCol;
            var channelStep = head.ChannelStep;

            // 所需枪头行、列数
            var tipChannelRow = tipChannel2DArray.GetLength(0);
            var tipChannelCol = tipChannel2DArray.GetLength(1);

            if (channelRow * channelCol > 1)
            {
                if (tipChannelCol > 1 && consumable.HoleStep.X != 0 && channelStep != consumable.HoleStep.X)
                    relation = channelStep / consumable.HoleStep.X;
                else if (tipChannelRow > 1 && consumable.HoleStep.Y != 0 && channelStep != consumable.HoleStep.Y)
                    relation = channelStep / consumable.HoleStep.Y;
            }

            return relation;
        }

        /// <summary>
        /// 获取枪头盒某行或某列所有未使用的枪头孔Index列表
        /// </summary>
        /// <param name="headUsedIndex"></param>
        /// <param name="tipChannel2DArray">枪头使用数量二维数组</param>
        /// <param name="consumable">枪头盒耗材</param>
        /// <param name="tipBoxUsedStatusArray"></param>
        /// <param name="rowCountOrColCount">所需枪头行数或列数</param>
        /// <param name="takeTipLeft2Right">取枪头方向</param>
        /// <returns></returns>
        public static List<int> GetTipBoxUnusedIndexList(int headUsedIndex, int[,] tipChannel2DArray, Consumable consumable, bool[] tipBoxUsedStatusArray, int rowCountOrColCount, bool takeTipLeft2Right)
        {
            var result = tipBoxUsedStatusArray.Select((p, i) => !p ? i : -1).Where(i => i != -1).ToList();
            // 按行取才需要判断
            if (!takeTipLeft2Right)
                result.Reverse();

            // 多通道移液头间距与枪头盒孔距比例，并按比例转为整形，获得间距倍数
            var multiple = Convert.ToInt16(TipHelper.MultiChannelStepAndConsumableStepRelation(headUsedIndex, tipChannel2DArray, consumable));
            // 如果倍数不为1且 剩余孔数大于所需枪头行数或列数，计算获取相应的枪头孔index
            if (multiple > 1 && result.Count > rowCountOrColCount)
            {
                var subList = new List<int>();
                var index = 0;
                for (var i = 0; i < rowCountOrColCount; i++)
                {
                    if (index < result.Count)
                        subList.Add(result[index]);

                    index += multiple;
                }
                result = subList;
            }

            return result.GetRange(0, rowCountOrColCount);
        }
    }
}
