using AutoLiquid_Library.Enum;
using AutoLiquid_Library.Exceptions;
using AutoLiquid_Library.Utils;
using AutoLiquid_GenScript_Single_Handling;
using ControlzEx.Standard;
using Serilog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Media3D;
using AutoLiquid_GenScript_Single_Handling.EntityCommon;
using AutoLiquid_GenScript_Single_Handling.EntityJson;

namespace AutoLiquid_GenScript_Single_Handling.Utils
{
    public class CmdHelper
    {
        // 通信窗体
        public static DAERun.FrmMain frmDAE = new DAERun.FrmMain(ParamsHelper.IO.IP, 4000, ParamsHelper.IO.Tcp);

        // 用户手动干预
        public static bool isManualPause;
        public static bool isManualStop;

        // 减速指令
        private static string SPEED_SLOW = "20 400 1";

        // 移液头左右轴列表
        private static List<string> HEAD_X_AXIS_LIST = new List<string> { "X", "X" };
        // 移液头前后轴列表
        private static List<string> HEAD_Y_AXIS_LIST = new List<string> { "Y", "Y" };
        // 移液头上下轴列表
        private static List<string> HEAD_Z_AXIS_LIST = new List<string> { "Z", "W" };
        // 移液头喷液轴列表
        public static List<string> HEAD_P_AXIS_LIST = new List<string> { "P", "Q" };
        // 移液头变距轴列表
        public static List<string> HEAD_Variable_AXIS_LIST = new List<string> { "W", "W" };

        // 盘位位置偏移值（包括盘面和退枪头卡位）
        public static Position offsetTemplate = new Position();

        // 移液头状态信息
        public static List<HeadStatus> headStatusList = new List<HeadStatus> { new HeadStatus(), new HeadStatus() };

        /**
         * 取枪头信息
         */
        // 取枪头孔位Index（一般用于指定盘位退枪头）
        private static HoleIndex takeTipPosIndex = new HoleIndex();
        // 取枪头信息
        public static TakeTipInfo TakeTipInfo = new TakeTipInfo();
        // 枪头盒位置偏移（适用于第二种枪头盒耗材）
        public static Position tipBoxOffset = new Position();
        // 枪头盒行走高度
        private static decimal tipBoxNormalHeight = 0;

        /**
         * 退枪头信息
         */
        // 退枪头盘位Index
        public static int? releaseTipTemplateIndex;


        // 指令超时时间
        private static long CMD_TIMEOUT = 20000;
        // 指令超时时间（离线调试模式）
        private static long CMD_TIMEOUT_TEST = 2000;

        /// <summary>
        /// 初始化机器
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// <param name="isFirstTimeInit">是否第一次复位</param>
        /// </summary>
        public static bool InitMachine(bool isNeedManualStop, bool isFirstTimeInit)
        {
            if (isNeedManualStop) LogHelper.Info((string)Application.Current.FindResource("Init"), "");

            /**
             * Z、P轴
             */
            var resetZ = ObjectUtils.GetMotionCmd(HEAD_Z_AXIS_LIST[0], EActType.I, "");
            resetZ += ParamsHelper.HeadList[1].Available ? "," + ObjectUtils.GetMotionCmd(HEAD_Z_AXIS_LIST[1], EActType.I, "") : "";
            var result1 = DoCmd(resetZ, isNeedManualStop);
            var resetP = ParamsHelper.HeadList[0].PAvailable ? ObjectUtils.GetMotionCmd(HEAD_P_AXIS_LIST[0], EActType.I, "") : "";
            resetP += ParamsHelper.HeadList[1].Available && ParamsHelper.HeadList[1].PAvailable ? "," + ObjectUtils.GetMotionCmd(HEAD_P_AXIS_LIST[1], EActType.I, "") : "";
            // 如果移液头1用Q轴推脱板
            if (ParamsHelper.HeadList[0].ReleaseTipUsePush && ParamsHelper.HeadList[0].ReleaseTipAxis == EAxis.Q)
                resetP += "," + ObjectUtils.GetMotionCmd(HEAD_P_AXIS_LIST[1], EActType.I, "");
            var result2 = DoCmd(resetP, isNeedManualStop);
            // 第一次复位，多做一次复位指令。避免还没建立通信，复位不了
            if (isFirstTimeInit)
            {
                result1 = DoCmd(resetZ, isNeedManualStop);
                result2 = DoCmd(resetP, isNeedManualStop);
            }

            /**
             * 变距轴
             */
            var resetVariable = ParamsHelper.HeadList[0].IsVariable && (ParamsHelper.HeadList[0].ChannelRow > 1 || ParamsHelper.HeadList[0].ChannelCol > 1) ? ObjectUtils.GetMotionCmd(HEAD_Variable_AXIS_LIST[0], EActType.I, "") : "";
            resetVariable += ParamsHelper.HeadList[1].Available && ParamsHelper.HeadList[1].IsVariable && (ParamsHelper.HeadList[1].ChannelRow > 1 || ParamsHelper.HeadList[1].ChannelCol > 1) ? "," + ObjectUtils.GetMotionCmd(HEAD_Variable_AXIS_LIST[1], EActType.I, "") : "";
            var result3 = DoCmd(resetVariable, isNeedManualStop);

            /**
             * X、Y轴
             */
            var resetXAndY = ParamsHelper.HeadList[0].YAvailable
                ? ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[0], EActType.I, "") + "," + ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[0], EActType.I, "")
                : ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[0], EActType.I, "");
            if (ParamsHelper.HeadList[1].Available)
            {
                if (ParamsHelper.HeadList[0].YAvailable)
                {
                    if (!HEAD_X_AXIS_LIST[0].Equals(HEAD_X_AXIS_LIST[1]))
                        resetXAndY += "," + ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[1], EActType.I, "");
                    if (!HEAD_Y_AXIS_LIST[0].Equals(HEAD_Y_AXIS_LIST[1]))
                        resetXAndY += "," + ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[1], EActType.I, "");
                }
                else
                {
                    if (!HEAD_X_AXIS_LIST[0].Equals(HEAD_X_AXIS_LIST[1]))
                        resetXAndY += "," + ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[1], EActType.I, "");
                }
            }
            var result4 = DoCmd(resetXAndY, isNeedManualStop);

            // 更新移液头状态信息
            headStatusList[0].Head = EHeadStatus.TipReleased;
            headStatusList[0].VolumeAbsorbMoreLeft = 0m;
            headStatusList[0].VolumeAirLeft = 0m;
            headStatusList[1].Head = EHeadStatus.TipReleased;
            headStatusList[1].VolumeAbsorbMoreLeft = 0m;
            headStatusList[1].VolumeAirLeft = 0m;

            return result1 && result2 && result3 && result4;
        }

        /// <summary>
        /// 初始化机器并把Y轴移动到方便摆放盘位置
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// <param name="isFirstTimeInit">是否第一次复位</param>
        /// </summary>
        public static bool InitMachineAndEasy2Put(bool isNeedManualStop, bool isFirstTimeInit)
        {
            if (isNeedManualStop) LogHelper.Info((string)Application.Current.FindResource("Init"), "");

            var result1 = InitMachine(isNeedManualStop, isFirstTimeInit);
            var result2 = true;
            // 如果y轴是盘移动，就移动到方便摆放盘位置
            if (!ParamsHelper.HeadList[0].YMoveWithHead && ParamsHelper.HeadList[0].YAvailable)
                result2 = Ya(0, ParamsHelper.CommonSettingList[0].PositionLaunchPlate.Y, EOffsetType.Template, isNeedManualStop);

            return result1 && result2;
        }

        /// <summary>
        /// 取枪头
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="consumableType">耗材类型</param>
        /// <param name="consumableTypeNext">下一个耗材类型（用于判断行走高度是走本耗材高度还是下一个耗材高度）</param>
        /// <param name="templateIndex">盘位index</param>
        /// <param name="holeIndex">枪头位置index</param>
        public static void TakeTip(int headUsedIndex, Consumable consumableType, Consumable consumableTypeNext, int templateIndex, HoleIndex holeIndex)
        {
            LogHelper.Info((string)Application.Current.FindResource("TakeTip"), "");

            /*
            * 获取所有可变值
            */
            var commonSetting = ParamsHelper.CommonSettingList[headUsedIndex];
            // 预取枪头高度
            var prepareTakeTipHeight = consumableType.PrepareTakeTipHeight + tipBoxOffset.Z;
            // 取枪头高度
            var takeTipHeight = consumableType.TakeTipHeight + tipBoxOffset.Z;
            // 行走高度
            tipBoxNormalHeight = consumableType.NormalHeight > consumableTypeNext.NormalHeight ? consumableTypeNext.NormalHeight + tipBoxOffset.Z : consumableType.NormalHeight + tipBoxOffset.Z;
            // 跨耗材行走高度复位
            var normalHeightReset = consumableType.NormalHeightReset;
            // 取枪头速度
            var takeTipSpeedCmd = consumableType.TakeTipSpeedCmd;
            // 重复取枪头次数
            var takeTipRepeatTime = consumableType.TakeTipRepeatTime;
            // 重复取枪头高度
            var takeTipRepeatHeight = consumableType.TakeTipRepeatHeight + tipBoxOffset.Z;
            // 取枪头后预提高度
            var takeTipAfterPrepareHeight = consumableType.TakeTipAfterPrepareHeight + tipBoxOffset.Z;
            // 预提高度后执行指令
            var takeTipAfterPrepareHeightCmd = consumableType.TakeTipAfterPrepareHeightCmd;
            // Z默认速度
            var defaultZSpeed = commonSetting.DefaultZSpeed;
            // Z速度百分比
            var zSpeedPercent = commonSetting.ZSpeedPercent;
            // 第1段提起速度
            var speedTipLift1 = commonSetting.SpeedTipLift1;
            // 第1段提起速度百分比
            var tipLift1SpeedPercent = commonSetting.TipLift1SpeedPercent;
            // 第1段提起高度
            var tipLift1Height = commonSetting.TipLift1Height;

            // 到孔位置（返回取枪头位置，以便用于返回取枪头位置退枪头）
            TakeTipInfo.Pos = GetHolePos(headUsedIndex, consumableType, templateIndex, holeIndex);
            TakeTipInfo.TemplateIndex = templateIndex;
            TakeTipInfo.Pos.X += tipBoxOffset.X;
            TakeTipInfo.Pos.Y += tipBoxOffset.Y;
            TakeTipInfo.Pos.Z = prepareTakeTipHeight;
            GotoHole(headUsedIndex, TakeTipInfo.Pos, consumableType, true);
            takeTipPosIndex = holeIndex;

            // 预取枪头高度
            if (prepareTakeTipHeight != 0)
            {
                GotoHeight(headUsedIndex, prepareTakeTipHeight, EOffsetType.Template, true);
                Thread.Sleep(200);
            }

            // 取枪头速度
            if (!takeTipSpeedCmd.Equals(""))
                SpeedSet(headUsedIndex, EAxis.Z, takeTipSpeedCmd, 100, true);

            // 取枪头高度
            GotoHeight(headUsedIndex, takeTipHeight, EOffsetType.Template, true);

            // 重复取枪头
            for (var i = 0; i < takeTipRepeatTime; i++)
            {
                // 重复取枪头高度
                GotoHeight(headUsedIndex, takeTipRepeatHeight, EOffsetType.Template, true);
                // 取枪头高度
                GotoHeight(headUsedIndex, takeTipHeight, EOffsetType.Template, true);
            }

            // 恢复Z轴速度
            if (!takeTipSpeedCmd.Equals(""))
                SpeedSet(headUsedIndex, EAxis.Z, defaultZSpeed, zSpeedPercent, true);

            // 更新移液头状态信息
            headStatusList[headUsedIndex].Head = EHeadStatus.TipTook;

            // 是否分段提起枪头
            if (tipLift1Height != 0)
            {
                SpeedSet(headUsedIndex, EAxis.Z, speedTipLift1, tipLift1SpeedPercent, true);
                GotoHeight(headUsedIndex, tipLift1Height, EOffsetType.Template, true);
                SpeedSet(headUsedIndex, EAxis.Z, defaultZSpeed, zSpeedPercent, true);
            }

            // 取枪头后预提高度
            if (takeTipAfterPrepareHeight != 0)
                GotoHeight(headUsedIndex, takeTipAfterPrepareHeight, EOffsetType.Template, true);
            // 预提高度后执行指令
            if (!takeTipAfterPrepareHeightCmd.Equals(""))
                DoCmdMulti(takeTipAfterPrepareHeightCmd.Trim(), true);

            // 行走高度
            GotoHeight(headUsedIndex, normalHeightReset ? 0 : tipBoxNormalHeight, EOffsetType.Template, true);
        }

        /// <summary>
        /// 吸液
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="consumableType">耗材类型</param>
        /// <param name="consumableTypeNext">下一个耗材类型（用于判断行走高度是走本耗材高度还是下一个耗材高度）</param>
        /// <param name="templateIndex">盘位 index</param>
        /// <param name="holeIndex">孔位</param>
        /// <param name="templateIndexNext">下一个盘位 index</param>
        /// <param name="seq">移液信息</param>
        /// <param name="liquidVolume">总吸液容积</param>
        /// <param name="isOne2More">是否一吸多喷</param>
        public static void AbsorbLiquid(int headUsedIndex, Consumable consumableType, Consumable consumableTypeNext, int templateIndex, HoleIndex holeIndex, int templateIndexNext, Seq seq, decimal liquidVolume, bool isOne2More)
        {
            LogHelper.Info((string)Application.Current.FindResource("AbsorbLiquid"), "");

            /*
           * 获取所有可变值
           */
            var commonSetting = ParamsHelper.CommonSettingList[headUsedIndex];
            // 行走高度
            var normalHeight = consumableType.NormalHeight > consumableTypeNext.NormalHeight ? consumableTypeNext.NormalHeight : consumableType.NormalHeight;
            // 跨耗材行走高度复位
            var normalHeightReset = consumableType.NormalHeightReset && templateIndex != templateIndexNext;
            // 混合高度
            var absorbMixingHeight = consumableType.AbsorbMixingHeight;
            // 吸液高度
            var liquidAbsorbHeight = consumableType.LiquidAbsorbHeight;
            // 吸液速度
            var absorbSpeed = consumableType.AbsorbSpeed;
            // 吸液后等待
            var liquidAbsorbDelay = consumableType.LiquidAbsorbDelay;
            // 吸液后提起速度
            var absorbHeight2NormalHeightSpeed = consumableType.AbsorbHeight2NormalHeightSpeed;
            // 吸液提起高度
            var absorbHeight2LiftingHeight = consumableType.AbsorbHeight2LiftingHeight;
            // 吸液提起后等待
            var liquidAbsorbDelayAfterLift = consumableType.LiquidAbsorbDelayAfterLift;
            // 喷液高度
            var liquidJetHeight = consumableType.LiquidJetHeight;
            // 喷液速度
            var jetSpeed = consumableType.JetSpeed;
            // 喷液后等待
            var liquidJetDelay = consumableType.LiquidJetDelay;
            // 靠壁高度
            var liquidAbsorbWallHeight = consumableType.LiquidJetWallHeight;
            // 靠壁偏移
            var liquidJetWallOffset = consumableType.LiquidJetWallOffset;
            // 量程
            var headLiquidRangeReal = Convert.ToDecimal((int)ParamsHelper.HeadList[headUsedIndex].HeadLiquidRange);
            // 吸液前吸空气量
            var absorbAirBefore = headLiquidRangeReal * commonSetting.AbsorbAirBeforePercent * 0.01m;
            // 喷液后喷空气等待时间
            var airDelayAfterJet = commonSetting.AirDelayAfterJet;
            // 吸液后吸空气量
            var absorbAirAfter = headLiquidRangeReal * commonSetting.AbsorbAirAfterPercent * 0.01m;
            // // 一吸多喷吸液后多吸体积
            // var absorbLiquidMoreOne2More = headLiquidRangeReal * commonSetting.AbsorbLiquidMoreOne2MorePercent * 0.01m;
            // // 一吸多喷吸液后喷出体积比例
            // var jetLiquidMoreOne2MoreScale = commonSetting.JetLiquidMoreOne2MoreScale;
            // 多点校准
            var multiCalibration = commonSetting.MultiCalibration;
            // 吸液后多吸体积（一吸多喷专用）
            var absorbLiquidMore = seq.SourceVolumeAbsorbMore > 0 ? seq.SourceVolumeAbsorbMore : isOne2More ? headLiquidRangeReal * commonSetting.AbsorbLiquidMorePercent * 0.01m : 0.0m;
            // 吸后反喷体积
            var reverseJetAfterAbsorb = seq.SourceVolumeReverseJet > 0 ? seq.SourceVolumeReverseJet : headLiquidRangeReal * commonSetting.ReverseJetAfterAbsorbPercent * 0.01m;
            // 是否梯度稀释
            var serialDilute = seq.SerialDilute;
            // 喷液体积补偿
            // var jetOffsetVolumeTotal = seq.JetOffsetList.Sum(jetOffset => jetOffset.VolumeOffset);
            // liquidVolume += jetOffsetVolumeTotal;
            // 喷液体积补偿（安全计算：允许 JetOffsetList 为 null 或为空）
            var jetOffsetVolumeTotal = seq?.JetOffsetList != null ? seq.JetOffsetList.Sum(jetOffset => jetOffset.VolumeOffset) : 0m;
            // 记录 jetOffset 的数量，后续做除法前必须检查 > 0
            var jetOffsetCount = seq?.JetOffsetList != null ? seq.JetOffsetList.Count : 0;
            liquidVolume += jetOffsetVolumeTotal;
            // 吸前混合速度
            var absorbMixingSpeed = consumableType.AbsorbMixingSpeed;
            // P默认速度
            var defaultPSpeed = commonSetting.DefaultPSpeed;
            // P速度百分比
            var pSpeedPercent = commonSetting.PSpeedPercent;
            // Z默认速度
            var defaultZSpeed = commonSetting.DefaultZSpeed;
            // Z速度百分比
            var zSpeedPercent = commonSetting.ZSpeedPercent;

            // 吸液前特殊指令
            if (!seq.CmdAbsorbBefore.Equals(""))
                MainWindow.mMainWindow.ParseCmdLine(seq.CmdAbsorbBefore, false);

            // 到孔位置
            GotoHole(headUsedIndex, consumableType, templateIndex, holeIndex, true);

            /**
             * 混合
             */
            var mixingVolume = seq.AbsorbMixingVolume;
            var mixingCount = seq.AbsorbMixingCount;
            var absorbMixingNeed = absorbMixingHeight > 0 && mixingVolume > 0 && mixingCount > 0;
            if (absorbMixingNeed)
            {
                // 吸液前吸空气量
                if (absorbAirBefore > 0)
                    Ps(headUsedIndex, absorbAirBefore, true);

                if (!absorbMixingSpeed.Equals(""))
                    SpeedSet(headUsedIndex, EAxis.P, absorbMixingSpeed, 100, true);

                // 混合高度
                GotoHeight(headUsedIndex, absorbMixingHeight, EOffsetType.Template, true);

                // 吸液混合前指令
                if (!seq.CmdAbsorbMixingBefore.Equals(""))
                    MainWindow.mMainWindow.ParseCmdLine(seq.CmdAbsorbMixingBefore, false);

                // 混合
                for (var i = 0; i < mixingCount; i++)
                {
                    // 吸液
                    Ps(headUsedIndex, mixingVolume, true);

                    // 喷液：最后一次喷液PA0
                    if (i == mixingCount - 1)
                    {
                        // 走到喷液高度
                        GotoHeight(headUsedIndex, liquidJetHeight, EOffsetType.Template, true);

                        Ps(headUsedIndex, mixingVolume * -1, true);
                        Thread.Sleep((int)(airDelayAfterJet * 1000));
                        Pa(headUsedIndex, 0, true);

                        // 吸液混合后指令
                        if (!seq.CmdAbsorbMixingAfter.Equals(""))
                            MainWindow.mMainWindow.ParseCmdLine(seq.CmdAbsorbMixingAfter, false);
                    }
                    else
                        Ps(headUsedIndex, mixingVolume * -1, true);
                }

                // 恢复用户设置的速度
                if (!absorbMixingSpeed.Equals(""))
                    SpeedSet(headUsedIndex, EAxis.P, defaultPSpeed, pSpeedPercent, true);
            }

            // 吸液前吸空气量
            if (absorbAirBefore > 0)
            {
                // 判断枪头是否已经前导气封，如果是，就不再需要多吸
                if (headStatusList[headUsedIndex].VolumeAirLeft == 0)
                {
                    // 如果吸液前需要混合，就先提高到行走高度，再吸空气
                    if (absorbMixingNeed)
                        GotoHeight(headUsedIndex, normalHeight, EOffsetType.Template, true);

                    Ps(headUsedIndex, absorbAirBefore, true);
                    headStatusList[headUsedIndex].VolumeAirLeft = absorbAirBefore;
                }
            }

            // 吸液高度
            GotoHeight(headUsedIndex, liquidAbsorbHeight, EOffsetType.Template, true);

            // 吸液速度
            if (!absorbSpeed.Equals(""))
                SpeedSet(headUsedIndex, EAxis.P, absorbSpeed, 100, true);

            // 吸液体积
            if (serialDilute)
            {
                var serialDiluteVolume = seq.VolumeEachList[0].Calibration;
                if (jetOffsetCount > 0)
                    serialDiluteVolume += jetOffsetVolumeTotal / jetOffsetCount;
                Ps(headUsedIndex, serialDiluteVolume, true);
            }
            else
            {
                if (liquidVolume > 0)
                {
                    if (multiCalibration.Available)
                        LogHelper.Info((string)Application.Current.FindResource("Prompt_Volume_Absorb_After_Multi_Calibration"), liquidVolume.ToString());
                    Ps(headUsedIndex, liquidVolume, true);
                }
            }

            /**
             * 如果是多吸体积
             */
            // 多吸体积
            if (absorbLiquidMore > 0 && liquidVolume > 0 && !serialDilute)
            {
                // 判断枪头是否已经多吸，如果是，就不再需要多吸
                if (headStatusList[headUsedIndex].VolumeAbsorbMoreLeft == 0)
                {
                    LogHelper.Info((string)Application.Current.FindResource("AbsorbLiquidMore"), "");
                    Ps(headUsedIndex, absorbLiquidMore, true);
                    headStatusList[headUsedIndex].VolumeAbsorbMoreLeft = absorbLiquidMore;
                }
            }
            // 吸后反喷体积
            if (reverseJetAfterAbsorb > 0 && liquidVolume > 0 && !serialDilute)
            {
                LogHelper.Info((string)Application.Current.FindResource("ReverseJetAfterAbsorb"), "");
                Ps(headUsedIndex, reverseJetAfterAbsorb, true);
            }
            // 吸液等待
            Thread.Sleep((int)(liquidAbsorbDelay * 1000));
            // 吸后反喷体积
            if (reverseJetAfterAbsorb > 0 && liquidVolume > 0 && !serialDilute)
                Ps(headUsedIndex, reverseJetAfterAbsorb * -1, true);

            // 恢复用户设置的速度
            if (!absorbSpeed.Equals(""))
                SpeedSet(headUsedIndex, EAxis.P, defaultPSpeed, pSpeedPercent, true);

            // 更新移液头状态信息
            headStatusList[headUsedIndex].SourceConsumableType = consumableType;
            headStatusList[headUsedIndex].SourceTemplateIndex = templateIndex;
            headStatusList[headUsedIndex].SourceHoleIndex = holeIndex;
            headStatusList[headUsedIndex].Head = EHeadStatus.Absorbed;

            // 吸液后提起速度
            if (!absorbHeight2NormalHeightSpeed.Equals(""))
                SpeedSet(headUsedIndex, EAxis.Z, absorbHeight2NormalHeightSpeed, 100, true);

            // 吸液靠壁
            if (seq.AbsorbWallList.Count > 0 && liquidAbsorbWallHeight > 0)
            {
                // 靠壁高度
                GotoHeight(headUsedIndex, liquidAbsorbWallHeight, EOffsetType.Template, true);

                // 提起后等待
                Thread.Sleep((int)(liquidAbsorbDelayAfterLift * 1000));

                /**
                 * 多向靠壁
                 */
                foreach (var absorbWall in seq.AbsorbWallList)
                {
                    // 左靠壁
                    if (absorbWall == EWall.Left)
                        WallAndReset(headUsedIndex, consumableType, templateIndex, holeIndex, HEAD_X_AXIS_LIST[headUsedIndex], SPEED_SLOW, "-" + liquidJetWallOffset.ToString("F2"));
                    // 右靠壁
                    else if (absorbWall == EWall.Right)
                        WallAndReset(headUsedIndex, consumableType, templateIndex, holeIndex, HEAD_X_AXIS_LIST[headUsedIndex], SPEED_SLOW, liquidJetWallOffset.ToString("F2"));
                    // 前靠壁
                    else if (absorbWall == EWall.Front)
                        WallAndReset(headUsedIndex, consumableType, templateIndex, holeIndex, HEAD_Y_AXIS_LIST[headUsedIndex], SPEED_SLOW, liquidJetWallOffset.ToString("F2"));
                    // 后靠壁
                    else if (absorbWall == EWall.Back)
                        WallAndReset(headUsedIndex, consumableType, templateIndex, holeIndex, HEAD_Y_AXIS_LIST[headUsedIndex], SPEED_SLOW, "-" + liquidJetWallOffset.ToString("F2"));
                }

                // 行走高度
                GotoHeight(headUsedIndex, normalHeightReset ? 0 : normalHeight, EOffsetType.Template, true);
            }
            else
            {
                // 吸液提起高度
                if (absorbHeight2LiftingHeight != 0)
                {
                    GotoHeight(headUsedIndex, absorbHeight2LiftingHeight, EOffsetType.Template, true);
                    // 提起后等待
                    Thread.Sleep((int)(liquidAbsorbDelayAfterLift * 1000));
                }

                // 恢复用户设置的速度
                if (!absorbHeight2NormalHeightSpeed.Equals(""))
                    SpeedSet(headUsedIndex, EAxis.Z, defaultZSpeed, zSpeedPercent, true);

                // 行走高度
                GotoHeight(headUsedIndex, normalHeightReset ? 0 : normalHeight, EOffsetType.Template, true);
            }

            // 恢复用户设置的速度
            if (!absorbHeight2NormalHeightSpeed.Equals(""))
                SpeedSet(headUsedIndex, EAxis.Z, defaultZSpeed, zSpeedPercent, true);

            // 吸液后吸空气量
            if (absorbAirAfter > 0)
                Ps(headUsedIndex, absorbAirAfter, true);

            // 吸液后特殊指令
            if (!seq.CmdAbsorbAfter.Equals(""))
                MainWindow.mMainWindow.ParseCmdLine(seq.CmdAbsorbAfter, false);
        }

        /// <summary>
        /// 喷液（一吸一喷，有混合和分段喷液功能，1vs1）
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="sourceConsumableType">源孔耗材类型</param>
        /// <param name="sourceTemplateIndex">源孔盘位Index</param>
        /// <param name="sourceHoleIndex">源盘孔位Index</param>
        /// <param name="consumableType">耗材类型</param>
        /// <param name="consumableTypeNext">下一个耗材类型</param>
        /// <param name="normalHeightRunPreReleaseTipHeight">行走高度是否走预退枪头高度（否：走下一个耗材类型行走高度）</param>
        /// <param name="templateIndex">靶盘盘位 index</param>
        /// <param name="holeIndex">靶盘孔位</param>
        /// <param name="seq">移液信息</param>
        /// <param name="liquidVolume">单个喷液体积</param>
        /// <param name="reJet2Source">多吸液体返回源孔喷出</param>
        /// <param name="pa0">是否PA0，默认为false</param>
        public static void JetSingleLiquid(int headUsedIndex, Consumable sourceConsumableType, int sourceTemplateIndex, HoleIndex sourceHoleIndex, Consumable consumableType, Consumable consumableTypeNext, bool normalHeightRunPreReleaseTipHeight, int templateIndex, HoleIndex holeIndex, Seq seq, decimal liquidVolume, bool reJet2Source, bool pa0 = false)
        {
            LogHelper.Info((string)Application.Current.FindResource("JetLiquid"), "");

            /*
             * 获取所有可变值
             */
            var commonSetting = ParamsHelper.CommonSettingList[headUsedIndex];
            // 量程
            var headLiquidRangeReal = Convert.ToDecimal((int)ParamsHelper.HeadList[headUsedIndex].HeadLiquidRange);
            // 行走高度
            decimal normalHeight;
            if (normalHeightRunPreReleaseTipHeight) // 走预退枪头高度
                normalHeight = commonSetting.PrepareReleaseTipPosList[0].Z;
            else // 走下一个耗材的行走高度
            {
                if (consumableTypeNext != null)
                    normalHeight = consumableType.NormalHeight > consumableTypeNext.NormalHeight
                        ? consumableTypeNext.NormalHeight
                        : consumableType.NormalHeight;
                else
                    normalHeight = consumableType.NormalHeight;
            }
            // 跨耗材行走高度复位
            var normalHeightReset = consumableType.NormalHeightReset;
            // 多点校准
            var multiCalibration = commonSetting.MultiCalibration;
            // 吸液后多吸体积
            var absorbLiquidMore = seq.SourceVolumeAbsorbMore > 0 ? seq.SourceVolumeAbsorbMore : 0.0m;
            // 是否需要复位PA0
            var needPa0 = pa0 || (headStatusList[headUsedIndex].VolumeAbsorbMoreLeft == 0m && absorbLiquidMore == 0);
            // 混合高度
            var jetMixingHeight = consumableType.JetMixingHeight;
            // 混合体积
            var mixingVolume = seq.JetMixingVolume;
            // 混合次数
            var mixingCount = seq.JetMixingCount;
            // 是否需要混合
            var jetMixingNeed = jetMixingHeight > 0 && mixingVolume > 0 && mixingCount > 0;
            // 喷液高度
            var liquidJetHeight = consumableType.LiquidJetHeight;
            // 喷液速度
            var jetSpeed = consumableType.JetSpeed;
            // 喷液后等待
            var liquidJetDelay = consumableType.LiquidJetDelay;
            // 第2段喷液体积
            var volumeJet2 = commonSetting.VolumeJet2;
            // 喷后混合速度
            var jetMixingSpeed = consumableType.JetMixingSpeed;
            // 靠壁高度
            var liquidJetWallHeight = consumableType.LiquidJetWallHeight;
            // 靠壁偏移
            var liquidJetWallOffset = consumableType.LiquidJetWallOffset;
            // 靠壁触发条件
            var liquidJetWallTrigger = consumableType.LiquidJetWallTrigger;
            // 靠壁喷液
            var wallJet = consumableType.WallJet;
            // 源盘吸液高度
            var sourceLiquidAbsorbHeight = sourceConsumableType.LiquidAbsorbHeight;
            // 喷液后提起速度
            var jetHeight2NormalHeightSpeed = consumableType.JetHeight2NormalHeightSpeed;
            // P默认速度
            var defaultPSpeed = commonSetting.DefaultPSpeed;
            // P速度百分比
            var pSpeedPercent = commonSetting.PSpeedPercent;
            // Z默认速度
            var defaultZSpeed = commonSetting.DefaultZSpeed;
            // Z速度百分比
            var zSpeedPercent = commonSetting.ZSpeedPercent;
            // 喷液后喷空气等待时间
            var airDelayAfterJet = commonSetting.AirDelayAfterJet;

            // 到孔位置
            GotoHole(headUsedIndex, consumableType, templateIndex, holeIndex, true);

            // 喷液前特殊指令
            if (!seq.CmdJetBefore.Equals(""))
                MainWindow.mMainWindow.ParseCmdLine(seq.CmdJetBefore, false);

            /**
           * 喷液
           */
            // 喷液体积补偿
            if (seq.JetOffsetList.Count > 0)
            {
                JetOffset jetOffset = seq.JetOffsetList.FirstOrDefault(p => p.PosIndex == 0);
                liquidVolume += jetOffset != null ? jetOffset.VolumeOffset : 0;
            }
            // 是否分段喷液
            var isJetSeparate = volumeJet2 > 0 && liquidVolume > volumeJet2;
            // 喷液速度
            if (!jetSpeed.Equals(""))
                SpeedSet(headUsedIndex, EAxis.P, jetSpeed, 100, true);

            // 校准后喷液体积
            if (multiCalibration.Available)
                LogHelper.Info((string)Application.Current.FindResource("Prompt_Volume_Jet_After_Multi_Calibration"), liquidVolume.ToString());
            // 先靠壁，再喷液
            if (wallJet)
            {
                WallFirstJetSecond(headUsedIndex, seq, consumableType, jetMixingNeed, needPa0, isJetSeparate, volumeJet2, templateIndex, holeIndex, liquidVolume, true);
            }
            // 先喷液，再靠壁
            else
            {
                // 喷液高度
                GotoHeight(headUsedIndex, liquidJetHeight, EOffsetType.Template, true);

                // 喷液
                JetLiquid(headUsedIndex, liquidVolume, jetMixingNeed, needPa0, isJetSeparate, volumeJet2, true);

                // 喷液等待
                Thread.Sleep((int)(liquidJetDelay * 1000));

                // 靠壁偏移(小于触发条件才偏移)
                if (liquidVolume <= liquidJetWallTrigger && seq.JetWallList.Count > 0 && liquidJetWallHeight > 0)
                {
                    // 靠壁高度
                    GotoHeight(headUsedIndex, liquidJetWallHeight, EOffsetType.Template, true);

                    /**
                     * 多向靠壁
                     */
                    foreach (var jetWall in seq.JetWallList)
                    {
                        // 左靠壁
                        if (jetWall == EWall.Left)
                            WallAndReset(headUsedIndex, consumableType, templateIndex, holeIndex, HEAD_X_AXIS_LIST[headUsedIndex], SPEED_SLOW, "-" + liquidJetWallOffset.ToString("F2"));
                        // 右靠壁
                        else if (jetWall == EWall.Right)
                            WallAndReset(headUsedIndex, consumableType, templateIndex, holeIndex, HEAD_X_AXIS_LIST[headUsedIndex], SPEED_SLOW, liquidJetWallOffset.ToString("F2"));
                        // 前靠壁
                        else if (jetWall == EWall.Front)
                            WallAndReset(headUsedIndex, consumableType, templateIndex, holeIndex, HEAD_Y_AXIS_LIST[headUsedIndex], SPEED_SLOW, liquidJetWallOffset.ToString("F2"));
                        // 后靠壁
                        else if (jetWall == EWall.Back)
                            WallAndReset(headUsedIndex, consumableType, templateIndex, holeIndex, HEAD_Y_AXIS_LIST[headUsedIndex], SPEED_SLOW, "-" + liquidJetWallOffset.ToString("F2"));
                    }
                }
            }

            // 恢复用户设置的速度
            if (!jetSpeed.Equals(""))
                SpeedSet(headUsedIndex, EAxis.P, defaultPSpeed, pSpeedPercent, true);

            /**
             * 混合
             */
            if (jetMixingNeed)
            {
                if (!jetMixingSpeed.Equals(""))
                    SpeedSet(headUsedIndex, EAxis.P, jetMixingSpeed, 100, true);

                // 混合高度
                GotoHeight(headUsedIndex, jetMixingHeight, EOffsetType.Template, true);

                // 喷液混合前指令
                if (!seq.CmdJetMixingBefore.Equals(""))
                    MainWindow.mMainWindow.ParseCmdLine(seq.CmdJetMixingBefore, false);

                // 混合
                for (var i = 0; i < mixingCount; i++)
                {
                    // 吸液
                    Ps(headUsedIndex, mixingVolume, true);

                    // 喷液
                    if (i == mixingCount - 1)
                    {
                        // 走到喷液高度
                        GotoHeight(headUsedIndex, liquidJetHeight, EOffsetType.Template, true);

                        Ps(headUsedIndex, mixingVolume * -1, true);
                        Thread.Sleep((int)(airDelayAfterJet * 1000));
                        if (needPa0)
                            Pa(headUsedIndex, 0, true);

                        // 喷液混合后指令
                        if (!seq.CmdJetMixingAfter.Equals(""))
                            MainWindow.mMainWindow.ParseCmdLine(seq.CmdJetMixingAfter, false);
                    }
                    else
                        Ps(headUsedIndex, mixingVolume * -1, true);
                }

                // 恢复用户设置的速度
                if (!jetMixingSpeed.Equals(""))
                    SpeedSet(headUsedIndex, EAxis.P, defaultPSpeed, pSpeedPercent, true);
            }

            // 喷液后提起速度
            if (!jetHeight2NormalHeightSpeed.Equals(""))
                SpeedSet(headUsedIndex, EAxis.Z, jetHeight2NormalHeightSpeed, 100, true);

            // 行走高度
            GotoHeight(headUsedIndex, normalHeightReset ? 0 : normalHeight, EOffsetType.Template, true);

            // 恢复用户设置的速度
            if (!jetHeight2NormalHeightSpeed.Equals(""))
                SpeedSet(headUsedIndex, EAxis.Z, defaultZSpeed, zSpeedPercent, true);

            // 喷液后特殊指令
            if (!seq.CmdJetAfter.Equals(""))
                MainWindow.mMainWindow.ParseCmdLine(seq.CmdJetAfter, false);

            // 是否有残余液体，有的话判断是否打回到源孔
            if (headStatusList[headUsedIndex].VolumeAbsorbMoreLeft > 0 && reJet2Source)
            {
                // // 到孔位置
                // GotoHole(headUsedIndex, sourceConsumableType, sourceTemplateIndex, sourceHoleIndex, true);
                // // 吸液高度
                // GotoHeight(headUsedIndex, sourceLiquidAbsorbHeight, EOffsetType.Template, true);
                // // 喷液
                // Pa(headUsedIndex, 0, true);
                // Thread.Sleep(200);
                // // 行走高度
                // GotoHeight(headUsedIndex, normalHeightReset ? 0 : normalHeight, EOffsetType.Template, true);

                ReJet2Source(headUsedIndex, seq, sourceConsumableType, sourceTemplateIndex, sourceHoleIndex, true);

                // 更新移液头状态信息
                headStatusList[headUsedIndex].Head = EHeadStatus.Jetted;
                headStatusList[headUsedIndex].VolumeAbsorbMoreLeft = 0m;
                headStatusList[headUsedIndex].VolumeAirLeft = 0m;
            }
        }

        /// <summary>
        /// 喷液（一吸多喷）
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="sourceConsumableType">源孔耗材类型</param>
        /// <param name="sourceTemplateIndex">源孔盘位Index</param>
        /// <param name="sourceHoleIndex">源盘孔位Index</param>
        /// <param name="consumableType">耗材类型</param>
        /// <param name="consumableTypeNext">下一个耗材类型</param>
        /// <param name="normalHeightRunPreReleaseTipHeight">行走高度是否走预退枪头高度（否：走下一个耗材类型行走高度）</param>
        /// <param name="templateIndexList">靶盘盘位List index</param>
        /// <param name="holeIndexList">靶盘孔位List index</param>
        /// <param name="seq">移液信息</param>
        /// <param name="liquidVolumeList">喷液容积List</param>
        /// <param name="reJet2Source">多吸液体返回源孔喷出</param>
        public static void JetMultiLiquid(int headUsedIndex, Consumable sourceConsumableType, int sourceTemplateIndex, HoleIndex sourceHoleIndex, Consumable consumableType, Consumable consumableTypeNext, bool normalHeightRunPreReleaseTipHeight, List<int> templateIndexList, List<HoleIndex> holeIndexList, Seq seq, List<decimal> liquidVolumeList, bool reJet2Source)
        {
            LogHelper.Info((string)Application.Current.FindResource("JetLiquid"), "");

            /*
             * 获取所有可变值
             */
            var commonSetting = ParamsHelper.CommonSettingList[headUsedIndex];
            // 盘内行走高度
            var normalHeightInner = consumableType.NormalHeight;
            // 盘外行走高度
            decimal normalHeightOuter;
            if (normalHeightRunPreReleaseTipHeight) // 走预退枪头高度
                normalHeightOuter = commonSetting.PrepareReleaseTipPosList[0].Z;
            else // 走下一个耗材的行走高度
            {
                if (consumableTypeNext != null)
                    normalHeightOuter = consumableType.NormalHeight > consumableTypeNext.NormalHeight
                        ? consumableTypeNext.NormalHeight
                        : consumableType.NormalHeight;
                else
                    normalHeightOuter = consumableType.NormalHeight;
            }
            // 跨耗材行走高度复位
            var normalHeightReset = consumableType.NormalHeightReset;
            // 混合高度
            var jetMixingHeight = consumableType.JetMixingHeight;
            // 混合体积
            var mixingVolume = seq.JetMixingVolume;
            // 混合次数
            var mixingCount = seq.JetMixingCount;
            // 是否梯度稀释
            var serialDilute = seq.SerialDilute;
            // 喷液高度
            var liquidJetHeight = consumableType.LiquidJetHeight;
            // 喷液速度
            var jetSpeed = consumableType.JetSpeed;
            // 喷液后等待
            var liquidJetDelay = consumableType.LiquidJetDelay;
            // 喷后混合速度
            var jetMixingSpeed = consumableType.JetMixingSpeed;
            // 靠壁高度
            var liquidJetWallHeight = consumableType.LiquidJetWallHeight;
            // 靠壁偏移
            var liquidJetWallOffset = consumableType.LiquidJetWallOffset;
            // 靠壁触发条件
            var liquidJetWallTrigger = consumableType.LiquidJetWallTrigger;
            // 靠壁喷液
            var wallJet = consumableType.WallJet;
            // 源盘行走高度
            var sourceNormalHeight = sourceConsumableType.NormalHeight;
            // 源盘吸液高度
            var sourceLiquidAbsorbHeight = sourceConsumableType.LiquidAbsorbHeight;
            // 量程
            var headLiquidRangeReal = Convert.ToDecimal((int)ParamsHelper.HeadList[headUsedIndex].HeadLiquidRange);
            // 吸液前吸空气量
            var absorbAirBefore = headLiquidRangeReal * commonSetting.AbsorbAirBeforePercent * 0.01m;
            // 吸液后吸空气量
            var absorbAirAfter = headLiquidRangeReal * commonSetting.AbsorbAirAfterPercent * 0.01m;
            // // 一吸多喷吸液后多吸体积
            // var absorbLiquidMoreOne2More = headLiquidRangeReal * commonSetting.AbsorbLiquidMoreOne2MorePercent * 0.01m;
            // // 一吸多喷吸液后喷出体积比例
            // var jetLiquidMoreScale = commonSetting.JetLiquidMoreOne2MoreScale;
            // 多点校准
            var multiCalibration = commonSetting.MultiCalibration;
            // 吸液后多吸体积
            var absorbLiquidMore = seq.SourceVolumeAbsorbMore > 0 ? seq.SourceVolumeAbsorbMore : headLiquidRangeReal * commonSetting.AbsorbLiquidMorePercent * 0.01m;
            // 是否需要复位PA0
            var needPa0 = headStatusList[headUsedIndex].VolumeAbsorbMoreLeft == 0m && absorbLiquidMore == 0;
            // 一吸多喷喷液后回吸体积
            var volumeBackAbsorb = seq.VolumeBackAbsorb;
            // 喷液后提起速度
            var jetHeight2NormalHeightSpeed = consumableType.JetHeight2NormalHeightSpeed;
            // P默认速度
            var defaultPSpeed = commonSetting.DefaultPSpeed;
            // P速度百分比
            var pSpeedPercent = commonSetting.PSpeedPercent;
            // Z默认速度
            var defaultZSpeed = commonSetting.DefaultZSpeed;
            // Z速度百分比
            var zSpeedPercent = commonSetting.ZSpeedPercent;
            // 喷液后喷空气等待时间
            var airDelayAfterJet = commonSetting.AirDelayAfterJet;

            // 靶孔数
            var holeCount = holeIndexList.Count;
            // 靶孔计数器
            var holeTick = 0;

            for (var i = 0; i < holeCount; i++)
            {
                // 靶盘位，判断是1靶多孔 or 多靶多孔
                var templateIndex = templateIndexList.Count > 1 ? templateIndexList[i] : templateIndexList[0];
                // 孔位
                var holeIndex = holeIndexList[i];
                // 体积
                var liquidVolume = liquidVolumeList.Count > 1 ? liquidVolumeList[i] : liquidVolumeList[0];
                // 喷液体积补偿
                JetOffset jetOffset = null;
                // 喷液体积补偿
                if (seq.JetOffsetList.Count > 0)
                {
                    jetOffset = seq.JetOffsetList.FirstOrDefault(p => p.PosIndex == i);
                    liquidVolume += jetOffset != null ? jetOffset.VolumeOffset : 0;
                }
                // 下一个孔体积
                var liquidVolumeNext = 0.0m;
                if (i != holeCount - 1)
                {
                    liquidVolumeNext = liquidVolumeList.Count > 1 ? liquidVolumeList[i + 1] : liquidVolumeList[0];
                    var jetOffsetNext = seq.JetOffsetList.FirstOrDefault(p => p.PosIndex == i + 1);
                    liquidVolumeNext += jetOffsetNext != null ? jetOffsetNext.VolumeOffset : 0;
                }

                // 喷液前特殊指令
                if (!seq.CmdJetBefore.Equals(""))
                    MainWindow.mMainWindow.ParseCmdLine(seq.CmdJetBefore, false);

                // 到孔位置
                GotoHole(headUsedIndex, consumableType, templateIndex, holeIndex, true);

                // 喷液速度
                if (!jetSpeed.Equals(""))
                    SpeedSet(headUsedIndex, EAxis.P, jetSpeed, 100, true);

                // 校准后喷液体积
                if (multiCalibration.Available)
                    LogHelper.Info((string)Application.Current.FindResource("Prompt_Volume_Jet_After_Multi_Calibration"), liquidVolume.ToString());
                // 先靠壁，再喷液
                if (wallJet)
                {
                    // 如果为第1个孔液体，喷出量 = 吸液后吸空气量 + 喷液体积
                    if (holeTick == 0 && holeCount > 1)
                        WallFirstJetSecond(headUsedIndex, seq, consumableType, false, false, false, 0, templateIndex, holeIndex, liquidVolume + absorbAirAfter, true);
                    // // 最后1个孔喷液，且没有多吸液体，直接PA0
                    // else if (holeTick == holeCount - 1 && (absorbLiquidMoreOne2More == 0 || (absorbLiquidMoreOne2More > 0 && jetLiquidMoreScale == 100)))
                    //     WallFirstJetSecond(headUsedIndex, seq, consumableType, false, true, false, 0, templateIndex, holeIndex, liquidVolume, true);
                    else
                        //     WallFirstJetSecond(headUsedIndex, seq, consumableType, false, false, false, 0, templateIndex, holeIndex, liquidVolume, true);
                        WallFirstJetSecond(headUsedIndex, seq, consumableType, false, false, false, 0, templateIndex, holeIndex, liquidVolume, true);
                }
                // 先喷液，再靠壁
                else
                {
                    // 喷液高度
                    GotoHeight(headUsedIndex, liquidJetHeight, EOffsetType.Template, true);

                    /**
                     * 移液体积
                     */
                    // 梯度稀释
                    if (serialDilute)
                    {
                        if (needPa0)
                            Pa(headUsedIndex, 0, true);
                        else
                            Ps(headUsedIndex, (liquidVolume + volumeBackAbsorb) * -1, true);
                        // 更新移液头状态信息
                        headStatusList[headUsedIndex].Head = EHeadStatus.Jetted;
                    }
                    else
                    {
                        // 如果为第1个孔液体，喷出量 = 吸液后吸空气量 + 喷液体积
                        if (holeTick == 0 && holeCount > 1)
                            Ps(headUsedIndex, (liquidVolume + absorbAirAfter) * -1, true);
                        // 最后1个孔喷液，且没有多吸液体，直接PA0
                        // else if (holeTick == holeCount - 1 && (absorbLiquidMoreOne2More == 0 || (absorbLiquidMoreOne2More > 0 && jetLiquidMoreScale == 100)))
                        else if (holeTick == holeCount - 1)
                        {
                            Ps(headUsedIndex, (liquidVolume + volumeBackAbsorb) * -1, true);
                            Thread.Sleep((int)(airDelayAfterJet * 1000));
                            if (needPa0)
                                headStatusList[headUsedIndex].VolumeAbsorbMoreLeft = 0m;
                            // 更新移液头状态信息
                            headStatusList[headUsedIndex].Head = EHeadStatus.Jetted;
                        }
                        else
                            Ps(headUsedIndex, (liquidVolume + volumeBackAbsorb) * -1, true);

                        // 回吸（最后一孔不用回吸）
                        if (holeTick != holeCount - 1 && volumeBackAbsorb != 0)
                            Ps(headUsedIndex, volumeBackAbsorb, true);
                    }

                    // 喷液等待
                    Thread.Sleep((int)(liquidJetDelay * 1000));

                    // 靠壁偏移(小于触发条件才偏移)
                    if (liquidVolume <= liquidJetWallTrigger && seq.JetWallList.Count > 0 && liquidJetWallHeight > 0)
                    {
                        GotoHeight(headUsedIndex, liquidJetWallHeight, EOffsetType.Template, true);

                        /**
                         * 多向靠壁
                         */
                        foreach (var jetWall in seq.JetWallList)
                        {
                            // 左靠壁
                            if (jetWall == EWall.Left)
                                WallAndReset(headUsedIndex, consumableType, templateIndex, holeIndex, HEAD_X_AXIS_LIST[headUsedIndex], SPEED_SLOW, "-" + liquidJetWallOffset.ToString("F2"));
                            // 右靠壁
                            if (jetWall == EWall.Right)
                                WallAndReset(headUsedIndex, consumableType, templateIndex, holeIndex, HEAD_X_AXIS_LIST[headUsedIndex], SPEED_SLOW, liquidJetWallOffset.ToString("F2"));
                            // 前靠壁
                            if (jetWall == EWall.Front)
                                WallAndReset(headUsedIndex, consumableType, templateIndex, holeIndex, HEAD_Y_AXIS_LIST[headUsedIndex], SPEED_SLOW, liquidJetWallOffset.ToString("F2"));
                            // 后靠壁
                            if (jetWall == EWall.Back)
                                WallAndReset(headUsedIndex, consumableType, templateIndex, holeIndex, HEAD_Y_AXIS_LIST[headUsedIndex], SPEED_SLOW, "-" + liquidJetWallOffset.ToString("F2"));
                        }
                    }
                }

                // 恢复用户设置的速度
                if (!jetSpeed.Equals(""))
                    SpeedSet(headUsedIndex, EAxis.P, defaultPSpeed, pSpeedPercent, true);

                /**
                 * 混合
                 */
                if (serialDilute)
                {
                    if (!jetMixingSpeed.Equals(""))
                        SpeedSet(headUsedIndex, EAxis.P, jetMixingSpeed, 100, true);

                    // 混合高度
                    GotoHeight(headUsedIndex, jetMixingHeight, EOffsetType.Template, true);

                    // 混合
                    for (var mixCountIndex = 0; mixCountIndex < mixingCount; mixCountIndex++)
                    {
                        // 吸液（= 混合体积 + 喷液补偿）
                        var volume = mixingVolume;
                        if (jetOffset != null) volume += jetOffset.VolumeOffset;
                        Ps(headUsedIndex, volume, true);
                        // 喷液
                        if (needPa0)
                            Pa(headUsedIndex, 0, true);
                        else
                            Ps(headUsedIndex, volume * -1, true);
                    }

                    // 下一次的喷液体积
                    if (liquidVolumeNext != 0)
                    {
                        // 吸液
                        Ps(headUsedIndex, liquidVolumeNext, true);
                    }

                    // 恢复用户设置的速度
                    if (!jetMixingSpeed.Equals(""))
                        SpeedSet(headUsedIndex, EAxis.P, defaultPSpeed, pSpeedPercent, true);
                }

                holeTick++;

                // 喷液后提起速度
                if (!jetHeight2NormalHeightSpeed.Equals(""))
                    SpeedSet(headUsedIndex, EAxis.Z, jetHeight2NormalHeightSpeed, 100, true);

                // 行走高度（如果非最后一次移液，就走盘内行走高度，否则走盘外行走高度）
                if (i != holeCount - 1)
                    GotoHeight(headUsedIndex, normalHeightInner, EOffsetType.Template, true);
                else
                    GotoHeight(headUsedIndex, normalHeightReset ? 0 : normalHeightOuter, EOffsetType.Template, true);

                // 恢复用户设置的速度
                if (!jetHeight2NormalHeightSpeed.Equals(""))
                    SpeedSet(headUsedIndex, EAxis.Z, defaultZSpeed, zSpeedPercent, true);

                // 喷液后特殊指令
                if (!seq.CmdJetAfter.Equals(""))
                    MainWindow.mMainWindow.ParseCmdLine(seq.CmdJetAfter, false);
            }

            // 是否有残余液体，有的话判断是否打回到源孔
            if (headStatusList[headUsedIndex].VolumeAbsorbMoreLeft > 0 && reJet2Source)
            {
                // LogHelper.Info((string)Application.Current.FindResource("ReJet2Source"), "");
                // // 到孔位置
                // GotoHole(headUsedIndex, sourceConsumableType, sourceTemplateIndex, sourceHoleIndex, true);
                // // 吸液高度
                // GotoHeight(headUsedIndex, sourceLiquidAbsorbHeight, EOffsetType.Template, true);
                // // 喷液
                // Pa(headUsedIndex, 0, true);
                // Thread.Sleep(200);
                // // 行走高度
                // GotoHeight(headUsedIndex, normalHeightReset ? 0 : normalHeightOuter, EOffsetType.Template, true);

                ReJet2Source(headUsedIndex, seq, sourceConsumableType, sourceTemplateIndex, sourceHoleIndex, true);

                // 更新移液头状态信息
                headStatusList[headUsedIndex].Head = EHeadStatus.Jetted;
                headStatusList[headUsedIndex].VolumeAbsorbMoreLeft = 0m;
                headStatusList[headUsedIndex].VolumeAirLeft = 0m;
            }
        }

        /// <summary>
        /// 先靠壁，再喷液
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="seq"></param>
        /// <param name="consumableType"></param>
        /// <param name="jetMixingNeed">喷液前是否混合</param>
        /// <param name="isPA0"></param>
        /// <param name="isJetSeparate"></param>
        /// <param name="volumeJet2"></param>
        /// <param name="templateIndex"></param>
        /// <param name="holeIndex"></param>
        /// <param name="liquidVolume"></param>
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        private static void WallFirstJetSecond(int headUsedIndex, Seq seq, Consumable consumableType, bool jetMixingNeed, bool isPA0, bool isJetSeparate, decimal volumeJet2, int templateIndex, HoleIndex holeIndex, decimal liquidVolume, bool isNeedManualStop)
        {
            /*
             * 获取所有可变值
             */
            var commonSetting = ParamsHelper.CommonSettingList[headUsedIndex];
            // 喷液后等待
            var liquidJetDelay = consumableType.LiquidJetDelay;
            // 靠壁高度
            var liquidJetWallHeight = consumableType.LiquidJetWallHeight;
            // 靠壁偏移
            var liquidJetWallOffset = consumableType.LiquidJetWallOffset;
            // X默认速度
            var defaultXSpeed = commonSetting.DefaultXSpeed;
            // X速度百分比
            var xSpeedPercent = commonSetting.XSpeedPercent;
            // Y默认速度
            var defaultYSpeed = commonSetting.DefaultYSpeed;
            // Y速度百分比
            var ySpeedPercent = commonSetting.YSpeedPercent;


            // 减速靠壁指令
            var slowAndWallAxisCmd = "";

            // 靠壁高度
            GotoHeight(headUsedIndex, liquidJetWallHeight, EOffsetType.Template, isNeedManualStop);

            // 靠壁逻辑：默认是在第1个靠壁方向喷液，其他方向只靠壁
            if (seq.JetWallList.Count > 0)
            {
                // 左靠壁
                if (seq.JetWallList[0] == EWall.Left)
                    slowAndWallAxisCmd = ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.F, SPEED_SLOW) + "," + ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.S, "-" + liquidJetWallOffset.ToString("F2"));
                // 右靠壁
                else if (seq.JetWallList[0] == EWall.Right)
                    slowAndWallAxisCmd = ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.F, SPEED_SLOW) + "," + ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.S, liquidJetWallOffset.ToString("F2"));
                // 前靠壁
                else if (seq.JetWallList[0] == EWall.Front)
                    slowAndWallAxisCmd = ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.F, SPEED_SLOW) + "," + ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.S, liquidJetWallOffset.ToString("F2"));
                // 后靠壁
                else if (seq.JetWallList[0] == EWall.Back)
                    slowAndWallAxisCmd = ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.F, SPEED_SLOW) + "," + ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.S, "-" + liquidJetWallOffset.ToString("F2"));

                // 减速靠壁
                DoCmd(slowAndWallAxisCmd, isNeedManualStop);
            }

            // 喷液
            JetLiquid(headUsedIndex, liquidVolume, jetMixingNeed, isPA0, isJetSeparate, volumeJet2, isNeedManualStop);

            // 是否有喷液后再次靠壁（从第2个靠壁方向开始）
            for (var i = 1; i < seq.JetWallList.Count; i++)
            {
                var jetWall = seq.JetWallList[i];
                // 靠壁喷液后恢复孔位置
                GotoHole(headUsedIndex, consumableType, templateIndex, holeIndex, isNeedManualStop);
                // 左靠壁
                if (jetWall == EWall.Left)
                    slowAndWallAxisCmd = ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.F, SPEED_SLOW) + "," + ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.S, "-" + liquidJetWallOffset.ToString("F2"));
                // 右靠壁
                else if (jetWall == EWall.Right)
                    slowAndWallAxisCmd = ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.F, SPEED_SLOW) + "," + ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.S, liquidJetWallOffset.ToString("F2"));
                // 前靠壁
                else if (jetWall == EWall.Front)
                    slowAndWallAxisCmd = ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.F, SPEED_SLOW) + "," + ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.S, liquidJetWallOffset.ToString("F2"));
                // 后靠壁
                else if (jetWall == EWall.Back)
                    slowAndWallAxisCmd = ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.F, SPEED_SLOW) + "," + ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.S, "-" + liquidJetWallOffset.ToString("F2"));

                // 减速靠壁
                DoCmd(slowAndWallAxisCmd, isNeedManualStop);
            }

            // 恢复用户设置的速度
            SpeedSet(headUsedIndex, EAxis.X, defaultXSpeed, xSpeedPercent, isNeedManualStop);
            SpeedSet(headUsedIndex, EAxis.Y, defaultYSpeed, ySpeedPercent, isNeedManualStop);

            // 喷液等待
            Thread.Sleep((int)(liquidJetDelay * 1000));

            // 靠壁喷液后恢复孔位置
            GotoHole(headUsedIndex, consumableType, templateIndex, holeIndex, isNeedManualStop);
        }

        /// <summary>
        /// 喷液
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="liquidVolume">液体量</param>
        /// <param name="jetMixingNeed">喷液前是否混合</param>
        /// <param name="isPA0">是否执行PA0，把液体全喷完</param>
        /// <param name="isJetSeparate">是否分段喷液</param>
        /// <param name="volumeJet2">第2段喷液</param>
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        private static void JetLiquid(int headUsedIndex, decimal liquidVolume, bool jetMixingNeed, bool isPA0, bool isJetSeparate, decimal volumeJet2, bool isNeedManualStop)
        {
            /*
            * 获取所有可变值
            */
            var commonSetting = ParamsHelper.CommonSettingList[headUsedIndex];
            // 第1段喷液速度
            var speedJet1 = commonSetting.SpeedJet1;
            // 第1段喷液速度百分比
            var jet1SpeedPercent = commonSetting.Jet1SpeedPercent;
            // 第2段喷液速度
            var speedJet2 = commonSetting.SpeedJet2;
            // 第2段喷液速度百分比
            var jet2SpeedPercent = commonSetting.Jet2SpeedPercent;
            // 两段喷液间停留时间（ms）
            var delayBetweenJet = commonSetting.DelayBetweenJet;
            // P默认速度
            var defaultPSpeed = commonSetting.DefaultPSpeed;
            // P速度百分比
            var pSpeedPercent = commonSetting.PSpeedPercent;
            // 喷液后喷空气等待时间
            var airDelayAfterJet = commonSetting.AirDelayAfterJet;

            // 分段喷液
            if (isJetSeparate)
            {
                // 第1段
                SpeedSet(headUsedIndex, EAxis.P, speedJet1, jet1SpeedPercent, isNeedManualStop);
                Ps(headUsedIndex, (liquidVolume - volumeJet2) * -1, isNeedManualStop);

                // 中间停留时间
                Thread.Sleep(delayBetweenJet);

                // 第2段
                SpeedSet(headUsedIndex, EAxis.P, speedJet2, jet2SpeedPercent, isNeedManualStop);
                if (isPA0)
                {
                    Pa(headUsedIndex, 0, isNeedManualStop);
                    // 更新移液头状态信息
                    headStatusList[headUsedIndex].Head = EHeadStatus.Jetted;
                }
                else
                {
                    Ps(headUsedIndex, volumeJet2 * -1, isNeedManualStop);
                }

                // 恢复P轴速度
                SpeedSet(headUsedIndex, EAxis.P, defaultPSpeed, pSpeedPercent, isNeedManualStop);
            }
            // 不分段喷液
            else
            {
                // 如果分段1设置了速度，就用该速度
                if (!speedJet1.Equals(""))
                {
                    SpeedSet(headUsedIndex, EAxis.P, speedJet1, jet1SpeedPercent, isNeedManualStop);
                }

                if (isPA0)
                {
                    Ps(headUsedIndex, liquidVolume * -1, isNeedManualStop);
                    if (!jetMixingNeed)
                    {
                        Thread.Sleep((int)(airDelayAfterJet * 1000));
                        Pa(headUsedIndex, 0, isNeedManualStop);
                        // 更新移液头状态信息
                        headStatusList[headUsedIndex].VolumeAbsorbMoreLeft = 0m;
                        headStatusList[headUsedIndex].VolumeAirLeft = 0m;
                    }
                    // 更新移液头状态信息
                    headStatusList[headUsedIndex].Head = EHeadStatus.Jetted;
                }
                else
                    Ps(headUsedIndex, liquidVolume * -1, isNeedManualStop);

                Thread.Sleep(100);

                // 恢复P轴速度
                SpeedSet(headUsedIndex, EAxis.P, defaultPSpeed, pSpeedPercent, isNeedManualStop);
            }
        }

        /// <summary>
        /// 靠壁，再复位
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="consumableType">耗材类型</param>
        /// <param name="templateIndex">盘位 index</param>
        /// <param name="holeIndex">孔位</param>
        /// <param name="slowSpeedCmd">减速指令</param>
        /// <param name="wallAxisCmd">靠壁轴指令</param>
        private static void WallAndReset(int headUsedIndex, Consumable consumableType, int templateIndex, HoleIndex holeIndex, string axis, string slowSpeedCmd, string wallAxisCmd)
        {
            /*
           * 获取所有可变值
           */
            var commonSetting = ParamsHelper.CommonSettingList[headUsedIndex];
            // X默认速度
            var defaultXSpeed = commonSetting.DefaultXSpeed;
            // X速度百分比
            var xSpeedPercent = commonSetting.XSpeedPercent;
            // Y默认速度
            var defaultYSpeed = commonSetting.DefaultYSpeed;
            // Y速度百分比
            var ySpeedPercent = commonSetting.YSpeedPercent;

            // 减速靠壁
            var slow = ObjectUtils.GetMotionCmd(axis, EActType.F, slowSpeedCmd);
            var wall = ObjectUtils.GetMotionCmd(axis, EActType.S, wallAxisCmd);
            DoCmd(slow + "," + wall, true);

            // 恢复用户设置的速度
            SpeedSet(headUsedIndex, EAxis.X, defaultXSpeed, xSpeedPercent, true);
            SpeedSet(headUsedIndex, EAxis.Y, defaultYSpeed, ySpeedPercent, true);

            // 靠壁后恢复孔位置
            Thread.Sleep(100);
            GotoHole(headUsedIndex, consumableType, templateIndex, holeIndex, true);
        }

        /// <summary>
        /// 液体喷回到源孔
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="seq">移液信息</param>
        /// <param name="sourceConsumableType">源盘耗材类型</param>
        /// <param name="sourceTemplateIndex">源盘盘位 index</param>
        /// <param name="sourceHoleIndex">源盘孔位</param>
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        public static bool ReJet2Source(int headUsedIndex, Seq seq, Consumable sourceConsumableType, int sourceTemplateIndex, HoleIndex sourceHoleIndex, bool isNeedManualStop)
        {
            LogHelper.Info((string)Application.Current.FindResource("ReJet2Source"), "");

            /*
             * 获取所有可变值
             */
            // 喷液高度
            var liquidAbsorbHeight = sourceConsumableType.LiquidAbsorbHeight;
            // 靠壁喷液
            var wallJet = sourceConsumableType.WallJet;

            // 行走高度
            GotoHeight(headUsedIndex, 0, EOffsetType.Template, isNeedManualStop);

            // 到孔位置
            GotoHole(headUsedIndex, sourceConsumableType, sourceTemplateIndex, sourceHoleIndex, isNeedManualStop);

            // 先靠壁，再喷液
            if (wallJet && seq.JetWallList.Count > 0)
                WallFirstJetSecond(headUsedIndex, seq, sourceConsumableType, false, true, false, 0, sourceTemplateIndex, sourceHoleIndex, 0, isNeedManualStop);
            else
            {
                // 吸液高度
                GotoHeight(headUsedIndex, liquidAbsorbHeight, EOffsetType.Template, isNeedManualStop);

                // 喷液
                JetLiquid(headUsedIndex, 0, false, true, false, 0, isNeedManualStop);
            }

            // 行走高度
            GotoHeight(headUsedIndex, 0, EOffsetType.Template, isNeedManualStop);

            return true;
        }

        /// <summary>
        /// 退枪头
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="isNeedZ0AfterRelease">退枪头后是否需要高度复位（即不同移液头之间切换，需要高度复位）</param>
        /// <param name="isMidwayOrZa0">是否中途退枪头 或 高度是否回零（如果是，先走高度0位，再走预退枪头位置，最后走预退枪头高度）</param>
        /// <param name="releaseTipPosIndex">退枪头位置Index</param>
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        public static void ReleaseTip(int headUsedIndex, bool isNeedZ0AfterRelease, bool isMidwayOrZa0, int releaseTipPosIndex, bool isNeedManualStop)
        {
            LogHelper.Info((string)Application.Current.FindResource("ReleaseTip"), "");

            /*
             * 获取所有可变值
             */
            var commonSetting = ParamsHelper.CommonSettingList[headUsedIndex];
            // 退枪头方式
            var releaseTipUsePush = ParamsHelper.HeadList[headUsedIndex].ReleaseTipUsePush;
            // 预退枪头位置
            Position prepareReleaseTipPos = new Position { X = commonSetting.PrepareReleaseTipPosList[releaseTipPosIndex].X, Y = commonSetting.PrepareReleaseTipPosList[releaseTipPosIndex].Y, Z = commonSetting.PrepareReleaseTipPosList[releaseTipPosIndex].Z };
            // 退枪头位置
            Position releaseTipPos = new Position { X = commonSetting.ReleaseTipPosList[releaseTipPosIndex].X, Y = commonSetting.ReleaseTipPosList[releaseTipPosIndex].Y, Z = commonSetting.ReleaseTipPosList[releaseTipPosIndex].Z };
            // 是否可变距
            var isVariableDistance = ParamsHelper.HeadList[headUsedIndex].IsVariable && (ParamsHelper.HeadList[headUsedIndex].ChannelRow > 1 || ParamsHelper.HeadList[headUsedIndex].ChannelCol > 1);
            // 预退枪头先走X轴
            var prepareReleaseTipAxisXGoFirst = commonSetting.PrepareReleaseTipAxisXGoFirst;
            // 预退枪头先走Y轴
            var prepareReleaseTipAxisYGoFirst = commonSetting.PrepareReleaseTipAxisYGoFirst;
            // 推脱板偏移
            var releaseTipOffset = commonSetting.ReleaseTipOffset;
            // 退枪头后指令
            var releaseTipAfterCmd = commonSetting.ReleaseTipAfterCmd;
            // 返回取枪头位置退枪头
            var releaseTipBack2TakePos = commonSetting.ReleaseTipBack2TakePos;
            // 退枪头前高度回零
            var releaseTipZa0Before = commonSetting.ReleaseTipZa0Before;
            // 退枪头后高度回零
            var releaseTipZa0After = commonSetting.ReleaseTipZa0After;
            // 退枪头速度指令
            var releaseTipSpeedCmd = commonSetting.ReleaseTipSpeedCmd;
            // P默认速度
            var defaultPSpeed = commonSetting.DefaultPSpeed;
            // P速度百分比
            var pSpeedPercent = commonSetting.PSpeedPercent;
            // Z默认速度
            var defaultZSpeed = commonSetting.DefaultZSpeed;
            // Z速度百分比
            var zSpeedPercent = commonSetting.ZSpeedPercent;

            // 返回取枪头位置退枪头
            if (releaseTipBack2TakePos)
            {
                prepareReleaseTipPos.X = TakeTipInfo.Pos.X;
                prepareReleaseTipPos.Y = TakeTipInfo.Pos.Y;
                releaseTipPos.X = TakeTipInfo.Pos.X;
                releaseTipPos.Y = TakeTipInfo.Pos.Y;
                releaseTipPos.Z = TakeTipInfo.Pos.Z;
            }
            // 指定盘位退枪头
            else if (releaseTipTemplateIndex != null)
            {
                var pos = GetHolePos(headUsedIndex, commonSetting.Consumables[0], (int)releaseTipTemplateIndex, takeTipPosIndex);
                prepareReleaseTipPos.X = pos.X;
                prepareReleaseTipPos.Y = pos.Y;
                releaseTipPos.X = pos.X;
                releaseTipPos.Y = pos.Y;
            }

            /**
             * 预退枪头规则：
             * ①中途退枪头或勾选了“退枪头前高度是否回零”：先走高度0位；再走预退枪头位置；最后走预退枪头高度
             * ②正常退枪头：先走预退枪头高度；再走预退枪头位置
             * @指定盘位退枪头：走到指定退枪头盘位的对应取枪头孔位退枪头
             */
            if (releaseTipZa0Before)
                isMidwayOrZa0 = true;
            // 高度0位 或者 预退枪头高度
            GotoHeight(headUsedIndex, isMidwayOrZa0 ? 0 : prepareReleaseTipPos.Z, EOffsetType.ReleaseTip, isNeedManualStop);
            /**
             * 变距
             * ①如果退枪头是卡扣，且设置了退枪头变距参数，就使用退枪头变距参数
             * ②否则，就使用枪头盒变距参数
             */
            if (isVariableDistance)
            {
                if (!releaseTipUsePush && commonSetting.ReleaseTipVariableDistanceStep != 0)
                    Wa(headUsedIndex, commonSetting.ReleaseTipVariableDistanceStep, isNeedManualStop);
                else
                    Wa(headUsedIndex, commonSetting.Consumables[0].VariableDistanceStep, isNeedManualStop);
            }


            // 预退枪头位置
            if (!ParamsHelper.HeadList[headUsedIndex].YMoveWhileReleaseTip) // 移液头1退枪头是否移动Y轴
                Xa(headUsedIndex, prepareReleaseTipPos.X, EOffsetType.ReleaseTip, isNeedManualStop);
            else
            {
                // Y轴不可用
                if (!ParamsHelper.HeadList[headUsedIndex].YAvailable)
                    Xa(headUsedIndex, prepareReleaseTipPos.X, EOffsetType.ReleaseTip, isNeedManualStop);
                else
                {
                    // 预退枪头先走X轴
                    if (prepareReleaseTipAxisXGoFirst)
                    {
                        Xa(headUsedIndex, prepareReleaseTipPos.X, EOffsetType.ReleaseTip, isNeedManualStop);
                        Ya(headUsedIndex, prepareReleaseTipPos.Y, EOffsetType.ReleaseTip, isNeedManualStop);
                    }
                    // 预退枪头先走Y轴
                    else if (prepareReleaseTipAxisYGoFirst)
                    {
                        Ya(headUsedIndex, prepareReleaseTipPos.Y, EOffsetType.ReleaseTip, isNeedManualStop);
                        Xa(headUsedIndex, prepareReleaseTipPos.X, EOffsetType.ReleaseTip, isNeedManualStop);
                    }
                    else
                    {
                        XaYa(headUsedIndex, prepareReleaseTipPos.X, prepareReleaseTipPos.Y, EOffsetType.ReleaseTip, isNeedManualStop);
                    }
                }
            }
            if (isMidwayOrZa0)
                // 预退枪头高度
                GotoHeight(headUsedIndex, prepareReleaseTipPos.Z, EOffsetType.ReleaseTip, isNeedManualStop);

            // 退枪头位置
            if (!ParamsHelper.HeadList[headUsedIndex].YMoveWhileReleaseTip) // 移液头1退枪头是否移动Y轴
                Xa(headUsedIndex, releaseTipPos.X, EOffsetType.ReleaseTip, isNeedManualStop);
            else
            {
                if (ParamsHelper.HeadList[headUsedIndex].YAvailable)
                    XaYa(headUsedIndex, releaseTipPos.X, releaseTipPos.Y, EOffsetType.ReleaseTip, isNeedManualStop);
                else
                    Xa(headUsedIndex, releaseTipPos.X, EOffsetType.ReleaseTip, isNeedManualStop);
            }
            // 退枪头速度
            var needSetReleaseTipSpeed = !releaseTipSpeedCmd.Equals("");
            // 退枪头速度（拉提方式）
            if (needSetReleaseTipSpeed && !ParamsHelper.HeadList[headUsedIndex].ReleaseTipUsePush)
                SpeedSet(headUsedIndex, EAxis.Z, releaseTipSpeedCmd, 100, isNeedManualStop);
            // 退枪头高度
            GotoHeight(headUsedIndex, releaseTipPos.Z, EOffsetType.ReleaseTip, isNeedManualStop);
            // 恢复Z速度（拉提方式）
            if (needSetReleaseTipSpeed && !ParamsHelper.HeadList[headUsedIndex].ReleaseTipUsePush)
                SpeedSet(headUsedIndex, EAxis.Z, defaultZSpeed, zSpeedPercent, isNeedManualStop);

            // 退枪头（推脱板方式）
            if (ParamsHelper.HeadList[headUsedIndex].ReleaseTipUsePush)
            {
                // 退枪头速度（推脱板方式）
                if (needSetReleaseTipSpeed)
                    SpeedSet(headUsedIndex, EAxis.P, releaseTipSpeedCmd, 100, isNeedManualStop);

                var releaseTipUsePushCount = ParamsHelper.HeadList[headUsedIndex].ReleaseTipUsePushCount;
                // 重复两次，避免枪头退不了
                for (var i = 0; i < releaseTipUsePushCount; i++)
                {
                    // 移液头1且推脱板为Q轴
                    if (headUsedIndex == 0 && ParamsHelper.HeadList[headUsedIndex].ReleaseTipAxis == EAxis.Q)
                    {
                        Pa(1, releaseTipOffset, isNeedManualStop);
                        Pa(1, 0, isNeedManualStop);
                    }
                    else
                    {
                        Pa(headUsedIndex, releaseTipOffset, isNeedManualStop);
                        Pa(headUsedIndex, 0, isNeedManualStop);
                    }
                }

                // 恢复Z速度（推脱板方式）
                if (needSetReleaseTipSpeed)
                    SpeedSet(headUsedIndex, EAxis.P, defaultPSpeed, pSpeedPercent, isNeedManualStop);
            }

            // 退枪头后指令
            if (!releaseTipAfterCmd.Equals(""))
                DoCmdMulti(releaseTipAfterCmd, isNeedManualStop);

            // 更新移液头状态信息
            headStatusList[headUsedIndex].Head = EHeadStatus.TipReleased;
            headStatusList[headUsedIndex].VolumeAbsorbMoreLeft = 0m;
            headStatusList[headUsedIndex].VolumeAirLeft = 0m;

            // 再次PA0
            Pa(headUsedIndex, 0, isNeedManualStop);

            // 高度复位
            if (isNeedZ0AfterRelease)
                GotoHeight(headUsedIndex, 0, EOffsetType.ReleaseTip, isNeedManualStop);
            // 退枪头后高度回零
            else if (releaseTipZa0After)
            {
                GotoHeight(headUsedIndex, 0, EOffsetType.ReleaseTip, isNeedManualStop);
                Zi(headUsedIndex, isNeedManualStop);
            }
            // 如果退枪头高度 > 枪头盒行走高度
            else if (releaseTipPos.Z > tipBoxNormalHeight)
                GotoHeight(headUsedIndex, tipBoxNormalHeight, EOffsetType.ReleaseTip, isNeedManualStop);
        }

        /// <summary>
        /// 到孔指定位置（X、Y轴）
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="consumableType">耗材类型</param>
        /// <param name="templateIndex">盘位index</param>
        /// <param name="holeIndex">孔index</param>
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// <returns>返回位置信息</returns>
        public static Position GotoHole(int headUsedIndex, Consumable consumableType, int templateIndex, HoleIndex holeIndex, bool isNeedManualStop)
        {
            // 获取孔位置
            var pos = GetHolePos(headUsedIndex, consumableType, templateIndex, holeIndex);

            GotoHole(headUsedIndex, pos, consumableType, isNeedManualStop);

            return pos;
        }

        /// <summary>
        /// 到孔指定位置（X、Y轴）
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="pos">位置</param>
        /// <param name="consumableType">耗材类型</param>
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// <returns>返回位置信息</returns>
        private static void GotoHole(int headUsedIndex, Position pos, Consumable consumableType, bool isNeedManualStop)
        {
            /*
              * 获取所有可变值
              */
            // 是否可变距
            var isVariableDistance = ParamsHelper.HeadList[headUsedIndex].IsVariable && (ParamsHelper.HeadList[headUsedIndex].ChannelRow > 1 || ParamsHelper.HeadList[headUsedIndex].ChannelCol > 1);
            // 变距步数
            var variableDistanceStep = consumableType.VariableDistanceStep;

            // 是否变距
            if (isVariableDistance)
            {
                // 变距与X、Y同时运动
                if (ParamsHelper.HeadList[headUsedIndex].VariableMoveSameTime)
                {
                    if (ParamsHelper.HeadList[headUsedIndex].YAvailable)
                        XaYaWa(headUsedIndex, pos.X, pos.Y, variableDistanceStep, EOffsetType.Template, isNeedManualStop);
                    else
                        XaWa(headUsedIndex, pos.X, variableDistanceStep, EOffsetType.Template, isNeedManualStop);
                }

                else
                {
                    if (ParamsHelper.HeadList[headUsedIndex].YAvailable)
                        XaYa(headUsedIndex, pos.X, pos.Y, EOffsetType.Template, isNeedManualStop);
                    else
                        Xa(headUsedIndex, pos.X, EOffsetType.Template, isNeedManualStop);
                    Wa(headUsedIndex, variableDistanceStep, isNeedManualStop);
                }
            }
            else
            {
                if (ParamsHelper.HeadList[headUsedIndex].YAvailable)
                    XaYa(headUsedIndex, pos.X, pos.Y, EOffsetType.Template, isNeedManualStop);
                else
                    Xa(headUsedIndex, pos.X, EOffsetType.Template, isNeedManualStop);
            }
        }

        /// <summary>
        /// 获取孔位置
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="consumableType">耗材类型</param>
        /// <param name="templateIndex">盘位index</param>
        /// <param name="holeIndex">孔index</param>
        /// <returns></returns>
        private static Position GetHolePos(int headUsedIndex, Consumable consumableType, int templateIndex, HoleIndex holeIndex)
        {
            // 位置信息
            var pos = new Position();

            // log holeIndex的信息
            LogHelper.Info("holeIndex信息：",
                           $"OriIndex={holeIndex.OriIndex}, " +
                           $"XHoleOffset={holeIndex.XHoleOffset}, " +
                           $"YHoleOffset={holeIndex.YHoleOffset}, " +
                           $"StepNotSameX={holeIndex.StepNotSameX}, " +
                           $"StepNotSameY={holeIndex.StepNotSameY}");

            /*
              * 获取所有可变值
              */
            // 移液头
            var head = ParamsHelper.HeadList[headUsedIndex];
            var commonSetting = ParamsHelper.CommonSettingList[headUsedIndex];
            // 首孔坐标
            var holeStartPos = consumableType.HoleStartPosList[templateIndex];
            // 盘位摆放（即A1方向）
            var a1Pos = commonSetting.A1Pos;
            // 取枪头方向：从左往右
            var takeTipLeft2Right = commonSetting.TakeTipLeft2Right || head.ChannelRow * head.ChannelCol == 1;
            // 行数
            var rowCount = consumableType.RowCount;
            // 列数
            var colCount = consumableType.ColCount;
            // 孔距
            var holeStep = consumableType.HoleStep;
            // 孔位Index
            var oriIndex = holeIndex.OriIndex;
            // 孔位偏移
            var xHoleOffset = holeIndex.XHoleOffset;
            var yHoleOffset = holeIndex.YHoleOffset;
            // 移液头通道间距与耗材间距是否不一致（如果不一致，先走移液头的偏移，再走耗材的偏移）
            var stepNotSameX = holeIndex.StepNotSameX;
            var stepNotSameY = holeIndex.StepNotSameY;

            var xIndex = 0;
            var yIndex = 0;

            if (oriIndex != 0)
            {
                // A1位置：左上
                if (a1Pos == EA1Pos.LeftTop)
                {
                    // 每列孔数
                    xIndex = oriIndex / rowCount;
                    yIndex = oriIndex % rowCount;
                }
                // A1位置：左下
                else
                {
                    // 每行孔数
                    xIndex = oriIndex % colCount;
                    yIndex = oriIndex / colCount;
                }
            }

            /**
             * x轴位置
             */
            decimal xa = holeStartPos.X + holeStep.X * xIndex;
            // 移液头通道间距与耗材间距不一致
            if (stepNotSameX)
            {
                // 如果耗材x间距为0，直接走移液头通道间距偏移
                if (xHoleOffset < 0)
                    xa += holeStep.X == 0 ? head.ChannelStep * xHoleOffset : (head.ChannelCol - 1) * head.ChannelStep * -1;
                else if (xHoleOffset > 0)
                    xa += holeStep.X == 0 ? head.ChannelStep * xHoleOffset : (head.ChannelCol - 1) * head.ChannelStep;
                else
                    // xa += takeTipLeft2Right ? (head.ChannelCol - 1) * head.ChannelStep * -1 : (head.ChannelCol - 1) * head.ChannelStep;
                    xa += takeTipLeft2Right ? (head.ChannelCol - 1) * head.ChannelStep * -1 : 0; // 2025-10-18修改
            }
            else
                xa += holeStep.X * xHoleOffset;
            /**
             * y轴位置
             */
            // 由于标准Y轴是盘推出，和多个移液头Y轴（Y轴由头移动）有差异
            var ya = holeStartPos.Y;
            if (ParamsHelper.HeadList[headUsedIndex].Available && ParamsHelper.HeadList[headUsedIndex].YMoveWithHead)
                ya = a1Pos == EA1Pos.LeftTop ? ya + holeStep.Y * yIndex : ya - holeStep.Y * yIndex;
            else
                ya = a1Pos == EA1Pos.LeftTop ? ya - holeStep.Y * yIndex : ya + holeStep.Y * yIndex;
            // 移液头通道间距与耗材间距不一致
            if (stepNotSameY)
            {
                // 如果耗材y间距为0，直接走移液头通道间距偏移
                if (yHoleOffset < 0)
                {
                    if (holeStep.Y == 0)
                        ya += head.ChannelStep * yHoleOffset;
                    else
                        // ya += a1Pos == EA1Pos.LeftTop ? (head.ChannelRow - 1) * head.ChannelStep * -1 : (head.ChannelRow - 1) * head.ChannelStep;
                        ya += (head.ChannelRow - 1) * head.ChannelStep * -1;
                }
                else if (yHoleOffset > 0)
                {
                    if (holeStep.Y == 0)
                        ya += head.ChannelStep * yHoleOffset;
                    else
                        // ya += a1Pos == EA1Pos.LeftTop ? (head.ChannelRow - 1) * head.ChannelStep : (head.ChannelRow - 1) * head.ChannelStep * -1;
                        ya += (head.ChannelRow - 1) * head.ChannelStep;
                }
            }
            else
                ya += holeStep.Y * yHoleOffset;

            // 赋值位置信息
            pos.X = xa;
            pos.Y = ya;

            return pos;
        }

        /// <summary>
        /// 到指定高度
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="height"></param>
        /// <param name="offsetType">偏移类型</param>
        /// <param name="isNeedManualStop"></param>
        public static bool GotoHeight(int headUsedIndex, decimal height, EOffsetType offsetType = EOffsetType.Template, bool isNeedManualStop = false)
        {
            var z = height;
            if (offsetType == EOffsetType.Template || offsetType == EOffsetType.ReleaseTip)
                z += offsetTemplate.Z;
            var cmd = ObjectUtils.GetMotionCmd(HEAD_Z_AXIS_LIST[headUsedIndex], EActType.A, z.ToString("F2"));
            return DoCmd(cmd, isNeedManualStop);
        }

        /// <summary>
        /// ZI
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="isNeedManualStop"></param>
        public static bool Zi(int headUsedIndex, bool isNeedManualStop = false)
        {
            var cmd = ObjectUtils.GetMotionCmd(HEAD_Z_AXIS_LIST[headUsedIndex], EActType.I, "");
            return DoCmd(cmd, isNeedManualStop);
        }

        /// <summary>
        /// XS
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="x"></param>
        /// <param name="offsetType">偏移类型</param>
        /// <param name="isNeedManualStop"></param>
        public static bool Xs(int headUsedIndex, decimal x, EOffsetType offsetType = EOffsetType.Template, bool isNeedManualStop = false)
        {
            if (offsetType == EOffsetType.Template || offsetType == EOffsetType.ReleaseTip)
                x += offsetTemplate.X;
            var cmd = ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.S, x.ToString("F2"));
            return DoCmd(cmd, isNeedManualStop);
        }

        /// <summary>
        /// XA
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="x"></param>
        /// <param name="offsetType">偏移类型</param>
        /// <param name="isNeedManualStop"></param>
        public static bool Xa(int headUsedIndex, decimal x, EOffsetType offsetType = EOffsetType.Template, bool isNeedManualStop = false)
        {
            if (offsetType == EOffsetType.Template || offsetType == EOffsetType.ReleaseTip)
                x += offsetTemplate.X;
            var cmd = ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.A, x.ToString("F2"));
            return DoCmd(cmd, isNeedManualStop);
        }

        /// <summary>
        /// XI
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="isNeedManualStop"></param>
        public static bool Xi(int headUsedIndex, bool isNeedManualStop = false)
        {
            var cmd = ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.I, "");
            return DoCmd(cmd, isNeedManualStop);
        }

        /// <summary>
        /// YA
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="y"></param>
        /// <param name="offsetType">偏移类型</param>
        /// <param name="isNeedManualStop"></param>
        public static bool Ya(int headUsedIndex, decimal y, EOffsetType offsetType = EOffsetType.Template, bool isNeedManualStop = false)
        {
            if (offsetType == EOffsetType.Template || offsetType == EOffsetType.ReleaseTip)
                y += offsetTemplate.Y;
            var cmd = ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.A, y.ToString("F2"));
            return DoCmd(cmd, isNeedManualStop);
        }

        /// <summary>
        /// YI
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="isNeedManualStop"></param>
        public static bool Yi(int headUsedIndex, bool isNeedManualStop = false)
        {
            var cmd = ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.I, "");
            return DoCmd(cmd, isNeedManualStop);
        }

        /// <summary>
        /// XA、YA
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="offsetType">偏移类型</param>
        /// <param name="isNeedManualStop"></param>
        public static bool XaYa(int headUsedIndex, decimal x, decimal y, EOffsetType offsetType = EOffsetType.Template, bool isNeedManualStop = false)
        {
            var cmd = "";
            var resultX = true;
            var resultY = true;
            if (offsetType == EOffsetType.Template || offsetType == EOffsetType.ReleaseTip)
            {
                x += offsetTemplate.X;
                y += offsetTemplate.Y;
            }

            if (ParamsHelper.HeadList[headUsedIndex].WalkingLogic == EWalkingLogic.XFirst)
            {
                cmd = ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.A, x.ToString("F2"));
                resultX = DoCmd(cmd, isNeedManualStop);
                cmd = ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.A, y.ToString("F2"));
                resultY = DoCmd(cmd, isNeedManualStop);
                return resultX && resultY;
            }
            if (ParamsHelper.HeadList[headUsedIndex].WalkingLogic == EWalkingLogic.YFirst)
            {
                cmd = ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.A, y.ToString("F2"));
                resultY = DoCmd(cmd, isNeedManualStop);
                cmd = ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.A, x.ToString("F2"));
                resultX = DoCmd(cmd, isNeedManualStop);
                return resultX && resultY;
            }
            return DoCmd(ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.A, x.ToString("F2")) + "," + ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.A, y.ToString("F2")), isNeedManualStop);
        }

        /// <summary>
        /// XA、YA、WA
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="offsetType">偏移类型</param>
        /// <param name="isNeedManualStop"></param>
        public static bool XaYaWa(int headUsedIndex, decimal x, decimal y, int w, EOffsetType offsetType = EOffsetType.Template, bool isNeedManualStop = false)
        {
            var cmd = "";
            var resultX = true;
            var resultY = true;
            if (offsetType == EOffsetType.Template || offsetType == EOffsetType.ReleaseTip)
            {
                x += offsetTemplate.X;
                y += offsetTemplate.Y;
            }

            if (ParamsHelper.HeadList[headUsedIndex].WalkingLogic == EWalkingLogic.XFirst)
            {
                cmd = ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.A, x.ToString("F2"));
                resultX = DoCmd(cmd, isNeedManualStop);
                cmd = ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.A, y.ToString("F2")) + "," + ObjectUtils.GetMotionCmd(HEAD_Variable_AXIS_LIST[headUsedIndex], EActType.A, w.ToString("F2"));
                resultY = DoCmd(cmd, isNeedManualStop);
                return resultX && resultY;
            }
            if (ParamsHelper.HeadList[headUsedIndex].WalkingLogic == EWalkingLogic.YFirst)
            {
                cmd = ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.A, y.ToString("F2"));
                resultY = DoCmd(cmd, isNeedManualStop);
                cmd = ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.A, x.ToString("F2")) + "," + ObjectUtils.GetMotionCmd(HEAD_Variable_AXIS_LIST[headUsedIndex], EActType.A, w.ToString("F2"));
                resultX = DoCmd(cmd, isNeedManualStop);
                return resultX && resultY;
            }
            return DoCmd(ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.A, x.ToString("F2")) + "," + ObjectUtils.GetMotionCmd(HEAD_Y_AXIS_LIST[headUsedIndex], EActType.A, y.ToString("F2")) + "," + ObjectUtils.GetMotionCmd(HEAD_Variable_AXIS_LIST[headUsedIndex], EActType.A, w.ToString("F2")), isNeedManualStop);
        }

        /// <summary>
        /// XA、WA
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="x"></param>
        /// <param name="offsetType">偏移类型</param>
        /// <param name="isNeedManualStop"></param>
        public static bool XaWa(int headUsedIndex, decimal x, int w, EOffsetType offsetType = EOffsetType.Template, bool isNeedManualStop = false)
        {
            if (offsetType == EOffsetType.Template || offsetType == EOffsetType.ReleaseTip)
                x += offsetTemplate.X;
            return DoCmd(ObjectUtils.GetMotionCmd(HEAD_X_AXIS_LIST[headUsedIndex], EActType.A, x.ToString("F2")) + "," + ObjectUtils.GetMotionCmd(HEAD_Variable_AXIS_LIST[headUsedIndex], EActType.A, w.ToString("F2")), isNeedManualStop);
        }

        /// <summary>
        /// PA
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="p"></param>
        /// <param name="isNeedManualStop"></param>
        public static bool Pa(int headUsedIndex, decimal p, bool isNeedManualStop = false)
        {
            if (ParamsHelper.HeadList[headUsedIndex].PAvailable)
            {
                var cmd = ObjectUtils.GetMotionCmd(HEAD_P_AXIS_LIST[headUsedIndex], EActType.A, p.ToString("F2"));
                return DoCmd(cmd, isNeedManualStop);
            }

            return true;
        }

        /// <summary>
        /// PS
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="p"></param>
        /// <param name="isNeedManualStop"></param>
        public static bool Ps(int headUsedIndex, decimal p, bool isNeedManualStop = false)
        {
            if (ParamsHelper.HeadList[headUsedIndex].PAvailable)
            {
                var cmd = ObjectUtils.GetMotionCmd(HEAD_P_AXIS_LIST[headUsedIndex], EActType.S, p.ToString("F2"));
                return DoCmd(cmd, isNeedManualStop);
            }

            return true;
        }

        /// <summary>
        /// PI
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="isNeedManualStop"></param>
        public static bool Pi(int headUsedIndex, bool isNeedManualStop = false)
        {
            if (ParamsHelper.HeadList[headUsedIndex].PAvailable)
            {
                var cmd = ObjectUtils.GetMotionCmd(HEAD_P_AXIS_LIST[headUsedIndex], EActType.I, "");
                return DoCmd(cmd, isNeedManualStop);
            }

            return true;
        }

        /// <summary>
        /// WA
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="w"></param>
        /// <param name="isNeedManualStop"></param>
        public static bool Wa(int headUsedIndex, int w, bool isNeedManualStop = false)
        {
            var cmd = ObjectUtils.GetMotionCmd(HEAD_Variable_AXIS_LIST[headUsedIndex], EActType.A, w.ToString());
            return DoCmd(cmd, isNeedManualStop);
        }

        /// <summary>
        /// WI
        /// </summary>
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="isNeedManualStop"></param>
        public static bool Wi(int headUsedIndex, bool isNeedManualStop = false)
        {
            var cmd = ObjectUtils.GetMotionCmd(HEAD_Variable_AXIS_LIST[headUsedIndex], EActType.I, "");
            return DoCmd(cmd, isNeedManualStop);
        }

        /// <summary>
        /// 设置轴速度（可复位该轴速度）
        /// <param name="headUsedIndex">移液头Index</param>
        /// <param name="axis">轴类型</param>
        /// <param name="speedCmd">速度指令（可不用填写轴前缀，如“FX”、“XF”等，且默认执行可刷写速度，如“XF”、“YF”等。如果为空，则单纯复位该轴速度）</param>
        /// <param name="speedPercent">速度百分比</param>
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static void SpeedSet(int headUsedIndex, EAxis axis, string speedCmd, decimal speedPercent, bool isNeedManualStop)
        {
            if (isNeedManualStop) LogHelper.Info((string)Application.Current.FindResource("SetSpeed"), "");

            var commonSetting = ParamsHelper.CommonSettingList[headUsedIndex];

            if (axis == EAxis.All)
            {
                SpeedSetSub(HEAD_X_AXIS_LIST[headUsedIndex], commonSetting.DefaultXSpeed, commonSetting.XSpeedPercent, isNeedManualStop);
                if (ParamsHelper.HeadList[headUsedIndex].YAvailable)
                    SpeedSetSub(HEAD_Y_AXIS_LIST[headUsedIndex], commonSetting.DefaultYSpeed, commonSetting.YSpeedPercent, isNeedManualStop);
                SpeedSetSub(HEAD_Z_AXIS_LIST[headUsedIndex], commonSetting.DefaultZSpeed, commonSetting.ZSpeedPercent, isNeedManualStop);
                if (ParamsHelper.HeadList[headUsedIndex].PAvailable)
                    SpeedSetSub(HEAD_P_AXIS_LIST[headUsedIndex], commonSetting.DefaultPSpeed, commonSetting.PSpeedPercent, isNeedManualStop);
            }
            else if (axis == EAxis.X)
            {
                SpeedSetSub(HEAD_X_AXIS_LIST[headUsedIndex], speedCmd, speedPercent, isNeedManualStop);
            }
            else if (axis == EAxis.Y && ParamsHelper.HeadList[headUsedIndex].YAvailable)
            {
                SpeedSetSub(HEAD_Y_AXIS_LIST[headUsedIndex], speedCmd, speedPercent, isNeedManualStop);
            }
            else if (axis == EAxis.Z)
            {
                SpeedSetSub(HEAD_Z_AXIS_LIST[headUsedIndex], speedCmd, speedPercent, isNeedManualStop);
            }
            else if (axis == EAxis.P)
            {
                if (ParamsHelper.HeadList[headUsedIndex].PAvailable)
                    SpeedSetSub(HEAD_P_AXIS_LIST[headUsedIndex], speedCmd, speedPercent, isNeedManualStop);
            }
        }

        /// <summary>
        /// 设置轴速度子方法
        /// </summary>
        /// <param name="axis">轴</param>
        /// <param name="speedCmd">速度指令</param>
        /// <param name="percent">速度百分比</param>
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        public static void SpeedSetSub(string axis, string speedCmd, decimal percent, bool isNeedManualStop)
        {
            speedCmd = speedCmd.Trim().ToUpper();
            if (speedCmd.Equals(""))
            {
                var initZeroCmd = ObjectUtils.GetMotionCmd(axis, EActType.F, "0");
                DoCmd(initZeroCmd, isNeedManualStop);
            }
            else
            {
                // 判断用户填写的速度指令是否包含前缀，如“FX”、“XF”等，有就自动去掉，只保留三个速度值，且默认执行可刷写速度，如“XF”、“YF”等
                if (speedCmd.Contains("F"))
                {
                    speedCmd = speedCmd.Substring(2).Trim();
                    if (speedCmd.Contains("@"))
                        speedCmd = speedCmd.Substring(0, speedCmd.IndexOf("@")).Trim();
                }
                var cmd = ObjectUtils.CalcSpeedCmdValueByPercent(axis, speedCmd, percent);
                DoCmd(cmd, isNeedManualStop);
            }
        }

        /// <summary>
        /// 开门状态和紧急停止功能状态查询
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static bool DoorAndEmergencyStopQuery(bool isNeedManualStop)
        {
            var result = true;

            if (ParamsHelper.IO.DoorAvailable)
                result = DoorQuery(isNeedManualStop);
            Thread.Sleep(20);
            if (ParamsHelper.IO.EmergencyStopAvailable)
                result = result && EmergencyStopQuery(isNeedManualStop);
            return result;
        }

        /// <summary>
        /// 紧急停止功能状态查询
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static bool EmergencyStopQuery(bool isNeedManualStop)
        {
            frmDAE.EmergencyStopCmd = ParamsHelper.IO.EmergencyStopCmd;
            LogHelper.Info((string)Application.Current.FindResource("Cmd"), frmDAE.EmergencyStopCmd);
            return DoCmd(frmDAE.EmergencyStopCmd, isNeedManualStop, false);
        }

        /// <summary>
        /// 紧急停止功能打开
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static void EmergencyStopOpen(bool isNeedManualStop)
        {
            LogHelper.Info((string)Application.Current.FindResource("EmergencyStopAvailable"), "");
            DoCmd(ParamsHelper.IO.EmergencyStopAvailableCmd, isNeedManualStop);
        }

        /// <summary>
        /// 紧急停止功能关闭
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static void EmergencyStopClose(bool isNeedManualStop)
        {
            LogHelper.Info((string)Application.Current.FindResource("EmergencyStopUnavailable"), "");
            DoCmd("A9 0", isNeedManualStop);
        }

        /// <summary>
        /// 开门状态查询
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static bool DoorQuery(bool isNeedManualStop)
        {
            frmDAE.DoorCmd = ParamsHelper.IO.DoorCmd;
            LogHelper.Info((string)Application.Current.FindResource("Cmd"), frmDAE.DoorCmd);
            return DoCmd(frmDAE.DoorCmd, isNeedManualStop, false);
        }

        /// <summary>
        /// 开门停止功能打开
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static void DoorStopOpen(bool isNeedManualStop)
        {
            LogHelper.Info((string)Application.Current.FindResource("DoorStopAvailable"), "");
            DoCmd(ParamsHelper.IO.DoorAvailableCmd, isNeedManualStop);
        }

        /// <summary>
        /// 开门停止功能关闭
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static void DoorStopClose(bool isNeedManualStop)
        {
            LogHelper.Info((string)Application.Current.FindResource("DoorStopUnavailable"), "");
            DoCmd("A9 0", isNeedManualStop);
        }

        /// <summary>
        /// 风扇打开
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static void FanOpen(bool isNeedManualStop)
        {
            LogHelper.Info((string)Application.Current.FindResource("FanAvailable"), "");
            DoCmd(ParamsHelper.IO.FanOpenCmd, isNeedManualStop);
        }

        /// <summary>
        /// 风扇关闭
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static void FanClose(bool isNeedManualStop)
        {
            LogHelper.Info((string)Application.Current.FindResource("FanUnavailable"), "");
            DoCmd(ParamsHelper.IO.FanCloseCmd, isNeedManualStop);
        }

        /// <summary>
        /// 照明打开
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static void LightOpen(bool isNeedManualStop)
        {
            LogHelper.Info((string)Application.Current.FindResource("LightAvailable"), "");
            DoCmd(ParamsHelper.IO.LightOpenCmd, isNeedManualStop);
        }

        /// <summary>
        /// 照明关闭
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static void LightClose(bool isNeedManualStop)
        {
            LogHelper.Info((string)Application.Current.FindResource("LightUnavailable"), "");
            DoCmd(ParamsHelper.IO.LightCloseCmd, isNeedManualStop);
        }

        /// <summary>
        /// 紫外灯打开
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static void UVOpen(bool isNeedManualStop)
        {
            LogHelper.Info((string)Application.Current.FindResource("UVAvailable"), "");
            DoCmd(ParamsHelper.IO.UVOpenCmd, isNeedManualStop);
        }

        /// <summary>
        /// 紫外灯关闭
        /// <param name="isNeedManualStop">是否需要手动停止功能</param>
        /// </summary>
        public static void UVClose(bool isNeedManualStop)
        {
            LogHelper.Info((string)Application.Current.FindResource("UVUnavailable"), "");
            DoCmd(ParamsHelper.IO.UVCloseCmd, isNeedManualStop);
        }

        /// <summary>
        /// 复位所有警示灯
        /// ao21亮绿色，运行中
        /// ao41亮黄色, 暂停
        //  ao31亮红灯, 出错
        /// </summary>
        /// <param name="temp"></param>
        /// <param name="isNeedManualStop"></param>
        /// <returns></returns>
        public static void ResetAllWarningLight(bool isNeedManualStop)
        {
            var cmd = ParamsHelper.IO.WarningLightGreenOpenCmd.Substring(0, 3) + "0,"
                      + ParamsHelper.IO.WarningLightYellowOpenCmd.Substring(0, 3) + "0,"
                      + ParamsHelper.IO.WarningLightRedOpenCmd.Substring(0, 3) + "0";
            DoCmd(cmd, isNeedManualStop);
        }

        /// <summary>
        /// 打开警示灯
        /// </summary>
        /// <param name="light"></param>
        /// <param name="isNeedManualStop"></param>
        public static void WarningLightOn(EWarningLight light, bool isNeedManualStop)
        {
            if (light == EWarningLight.Green)
                DoCmd(ParamsHelper.IO.WarningLightGreenOpenCmd, isNeedManualStop);
            else if (light == EWarningLight.Yellow)
                DoCmd(ParamsHelper.IO.WarningLightYellowOpenCmd, isNeedManualStop);
            else if (light == EWarningLight.Red)
                DoCmd(ParamsHelper.IO.WarningLightRedOpenCmd, isNeedManualStop);
        }

        /// <summary>
        /// 关闭警示灯
        /// </summary>
        /// <param name="light"></param>
        /// <param name="isNeedManualStop"></param>
        public static void WarningLightOff(EWarningLight light, bool isNeedManualStop)
        {
            if (light == EWarningLight.Green)
                DoCmd(ParamsHelper.IO.WarningLightGreenOpenCmd.Substring(0, 3) + "0", isNeedManualStop);
            else if (light == EWarningLight.Yellow)
                DoCmd(ParamsHelper.IO.WarningLightYellowOpenCmd.Substring(0, 3) + "0", isNeedManualStop);
            else if (light == EWarningLight.Red)
                DoCmd(ParamsHelper.IO.WarningLightRedOpenCmd.Substring(0, 3) + "0", isNeedManualStop);
        }

        /// <summary>
        /// 检测试管是否存在
        /// </summary>
        /// <param name="isNeedManualStop"></param>
        public static void CheckTubeExist(bool isNeedManualStop)
        {
            LogHelper.Info((string)Application.Current.FindResource("CheckTubeExist"), "");
            DoCmd("AA1", isNeedManualStop);
        }

        /// <summary>
        /// 查询机器设备id
        /// </summary>
        /// <returns></returns>
        public static bool QueryDeviceId()
        {
            return true;
            //            return DoCmd("", false);
        }

        /// <summary>
        /// 执行指令（指令执行入口，免除了繁琐的ManualStop调用）
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="isNeedManualStop">是否需要手动停止功能（用于仓门）</param>
        /// <param name="intercept">如果启用了门，是否需要拦截操作</param>
        /// <param name="waitAnswer">是否等待下位机执行完毕应答</param>
        /// <returns></returns>
        public static bool DoCmd(string cmd, bool isNeedManualStop, bool intercept = true, bool waitAnswer = true)
        {
            var cmdTimeOut = CMD_TIMEOUT;

            if (cmd.Equals(""))
                return true;

            LogHelper.Info((string)Application.Current.FindResource("Cmd"), cmd);

            if (!CmdHelper.frmDAE.isCanNetSuccess || frmDAE.socketSend == null || (ParamsHelper.IO.Tcp && !frmDAE.socketSend.Connected))
            {
                Thread.Sleep(10);
                if (isNeedManualStop)
                    ManualStop(intercept);
                return true;
            }


            // 已经过去的时间
            var timePassedMS = 0;
            var result = false;
            new Thread(() =>
            {
                result = CmdHelper.frmDAE.DoCmd(cmd, waitAnswer);
            }).Start();

            while (timePassedMS < cmdTimeOut)
            {
                // 指令被拦截了
                if (MainWindow.mIsEmergencyStopIntercept && intercept)
                {
                    // 要排除A开头的指令
                    if (!cmd.ToUpper().StartsWith("A"))
                        MainWindow.mEmergencyStopLastCmd = cmd;
                    break;
                }

                if (result)
                    break;

                Thread.Sleep(10);
                timePassedMS += 10;
            }

            // 超时
            if (!MainWindow.mIsEmergencyStopIntercept)
            {
                if (!result)
                {
                    // 打开红指示灯
                    if (ParamsHelper.IO.WarningLightAvailable)
                        CmdHelper.WarningLightOn(EWarningLight.Red, false);

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (MessageBox.Show((string)Application.Current.FindResource("Prompt_Cmd_TimeOut_1") + cmd + (string)Application.Current.FindResource("Prompt_Cmd_TimeOut_2"),
                                (string)Application.Current.FindResource("Prompt"), MessageBoxButton.YesNo, MessageBoxImage.Warning) ==
                            MessageBoxResult.No)
                        {
                            // 关闭红指示灯
                            if (ParamsHelper.IO.WarningLightAvailable)
                                CmdHelper.WarningLightOff(EWarningLight.Red, false);
                            throw new ManualStopException();
                        }
                    });

                    // 关闭红指示灯
                    if (ParamsHelper.IO.WarningLightAvailable)
                        CmdHelper.WarningLightOff(EWarningLight.Red, false);
                }
            }

            if (isNeedManualStop)
                ManualStop(intercept);

            return result;
        }

        /// <summary>
        /// 执行多指令（分号分隔）
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="isNeedManualStop">是否需要手动停止功能（用于仓门）</param>
        private static void DoCmdMulti(string cmd, bool isNeedManualStop)
        {
            // 拆分分号
            string[] semiCmd = cmd.Split(new[] { ';', '；' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var semi in semiCmd)
            {
                if (isNeedManualStop) LogHelper.Info((string)Application.Current.FindResource("Cmd"), semi);
                DoCmd(semi, isNeedManualStop);
            }
        }

        /// <summary>
        /// 手动停止
        /// </summary>
        /// <param name="intercept">如果启用了门，是否需要拦截操作</param>
        public static void ManualStop(bool intercept = true)
        {
            while (isManualPause)
            {
                if (isManualStop)
                {
                    throw new ManualStopException();
                }

                try
                {
                    Thread.Sleep(10);
                }
                catch (Exception e)
                {
                }
            }

            if (intercept)
            {
                // 恢复执行上一条被拦截的指令
                if (MainWindow.mIsEmergencyStopClosedAgain)
                {
                    MainWindow.mIsEmergencyStopClosedAgain = false;
                    DoCmd(MainWindow.mEmergencyStopLastCmd, true, true, true);
                }
            }
        }
    }
}
