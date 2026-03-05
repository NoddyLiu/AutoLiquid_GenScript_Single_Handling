using AutoLiquid_Library.Enum;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace AutoLiquid_Library.Utils
{
    /// <summary>
    /// 对象通用工具类
    /// </summary>
    public static class ObjectUtils
    {
        /// <summary>
        /// 深拷贝（不会共用对象的引用）
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="other"></param>
        /// <returns></returns>
        public static T DeepCopy<T>(T other)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter formatter = new BinaryFormatter();
                formatter.Context = new StreamingContext(StreamingContextStates.Clone);
                formatter.Serialize(ms, other);
                ms.Position = 0;
                return (T)formatter.Deserialize(ms);
            }
        }

        /// <summary>
        /// 只保留字符串的字母数字，并转为大写
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string GetAlphanumericToUpperOnly(string str)
        {
            return Regex.Replace(str, "[^a-zA-Z0-9]", "").Replace(".", "").ToUpper();
        }

        /// <summary>
        /// 根据标记获取特殊指令
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="cmdTag"></param>
        /// <param name="isTxtLink"></param>
        /// <returns></returns>
        public static string GetCmdAccordTag(string cmd, string cmdTag, bool isTxtLink)
        {
            // cmd是否有多行
            string[] lineCmd = cmd.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var t in lineCmd)
            {
                var regex = ";(?![^()]*\\))";
                string[] cmdOneLines = Regex.Split(t, regex);
                // string[] cmdOneLines = t.Split(new[] { ';', '；' }, StringSplitOptions.RemoveEmptyEntries);

                if (cmdOneLines.Length == 1)
                {
                    var cmdSingle = cmdOneLines[0].ToLower();
                    if (cmdSingle.Contains(cmdTag.ToLower()))
                    {
                        if (isTxtLink)
                            return cmdSingle.Split(new[] { ':', '：' }, StringSplitOptions.RemoveEmptyEntries).ElementAt(1);
                        if ((cmdSingle.Contains("(") || cmdSingle.Contains("（")) && !cmdSingle.Contains(ConstantsUtils.JetOffsetCmd.ToLower()))
                            return cmdSingle.Replace(cmdTag.ToLower(), "").Replace("(", "").Replace("（", "").Replace(")", "").Replace("）", "");
                        return cmdSingle.Split(new[] { ';', '；' }, StringSplitOptions.RemoveEmptyEntries).ElementAt(0).Split(new[] { ':', '：' }, StringSplitOptions.RemoveEmptyEntries).ElementAt(1);
                    }
                }
                else if (cmdOneLines.Length > 1)
                {
                    for (var i = 0; i < cmdOneLines.Length; i++)
                    {
                        var cmdSingle = cmdOneLines[i].ToLower();
                        if (cmdSingle.Contains(cmdTag.ToLower()))
                        {
                            if ((cmdSingle.Contains("(") || cmdSingle.Contains("（")) && !cmdSingle.Contains(ConstantsUtils.JetOffsetCmd.ToLower()))
                                return cmdSingle.Replace(cmdTag.ToLower(), "").Replace("(", "").Replace("（", "").Replace(")", "").Replace("）", "");
                            return cmdSingle.Split(new[] { ':', '：' }, StringSplitOptions.RemoveEmptyEntries).ElementAt(1);
                        }
                    }
                }
            }
            return "";
        }

        /// <summary>
        /// 根据百分比计算速度指令
        /// </summary>
        /// <param name="axis">轴</param>
        /// <param name="speedCmd">速度指令</param>
        /// <param name="percent">百分比</param>
        /// <returns></returns>
        public static string CalcSpeedCmdValueByPercent(string axis, string speedCmd, decimal percent)
        {
            if (speedCmd.Equals(""))
                return "";

            // 输出指令
            var outCmdStr = "";
            if (percent >= 10.0m)
            {
                // 以空格分隔数值
                var valueArray = speedCmd.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                // 重组指令
                foreach (var value in valueArray)
                {
                    outCmdStr += " " + Convert.ToInt32(Int32.Parse(value) * percent * 0.01m);
                }
            }

            return GetMotionCmd(axis, EActType.F, outCmdStr);
        }

        /// <summary>
        /// 根据子list内容个数拆分列表
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        public static List<List<T>> SplitList<T>(List<T> source, int size)
        {
            var list = new List<List<T>>();

            for (int i = 0; i < source.Count; i += size)
            {
                if (i + size < source.Count)
                    list.Add(source.GetRange(i, size));
                // 最后的超出范围，要截取
                else
                {
                    var lastSize = source.Count - i;
                    list.Add(source.GetRange(i, lastSize));
                }

            }

            return list;
        }

        /// <summary>
        /// 二维数组数值填充
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="arr"></param>
        /// <param name="value"></param>
        public static void Fill2DArray<T>(T[,] arr, T value)
        {
            int numRows = arr.GetLength(0);
            int numCols = arr.GetLength(1);

            for (int i = 0; i < numRows; ++i)
            {
                for (int j = 0; j < numCols; ++j)
                {
                    arr[i, j] = value;
                }
            }
        }

        /// <summary>
        /// 检查是否包含特殊指令
        /// </summary>
        /// <param name="cmd"></param>
        /// <returns></returns>
        public static bool CheckCmdExist(string cmd)
        {
            // 剔除移液头特殊指令
            if (cmd.ToLower().Contains(ConstantsUtils.Head2Cmd.ToLower()))
                return true;

            // 剔除是否最后退枪头特殊指令
            if (cmd.ToLower().Contains(ConstantsUtils.NoReleaseTipCmd.ToLower()))
                return true;

            // 剔除是否强制退枪头特殊指令
            if (cmd.ToLower().Contains(ConstantsUtils.ReleaseTipCmd.ToLower()))
                return true;

            // 剔除吸液、喷液前后特殊指令
            if (cmd.ToLower().Contains(ConstantsUtils.AbsorbBeforeCmd.ToLower()) || cmd.ToLower().Contains(ConstantsUtils.AbsorbAfterCmd.ToLower())
                                                                                                           || cmd.ToLower().Contains(ConstantsUtils.JetBeforeCmd.ToLower()) || cmd.ToLower().Contains(ConstantsUtils.JetAfterCmd.ToLower()))
                return true;

            // 剔除指定盘位退枪头指令
            if (cmd.ToLower().Contains(ConstantsUtils.ReleaseTipTemplateCmd.ToLower()))
                return true;

            // 剔除第2种枪头盒相对第1种枪头盒位置偏移指令
            if (cmd.ToLower().Contains(ConstantsUtils.TipBoxOffsetXCmd.ToLower())
                || cmd.ToLower().Contains(ConstantsUtils.TipBoxOffsetYCmd.ToLower())
                || cmd.ToLower().Contains(ConstantsUtils.TipBoxOffsetZCmd.ToLower()))
                return true;

            // 剔除指定孔位喷液体积补偿指令
            if (cmd.ToLower().Contains(ConstantsUtils.JetOffsetCmd.ToLower()))
                return true;

            // 剔除重复执行Excel表格
            if (cmd.ToLower().Contains(ConstantsUtils.RepeatCmd.ToLower()))
                return true;

            // 剔除混合前特殊指令
            if (cmd.ToLower().Contains(ConstantsUtils.AbsorbMixingBeforeCmd.ToLower()))
                return true;
            if (cmd.ToLower().Contains(ConstantsUtils.JetMixingBeforeCmd.ToLower()))
                return true;

            // 剔除混合后特殊指令
            if (cmd.ToLower().Contains(ConstantsUtils.AbsorbMixingAfterCmd.ToLower()))
                return true;
            if (cmd.ToLower().Contains(ConstantsUtils.JetMixingAfterCmd.ToLower()))
                return true;

            // 剔除一吸多喷喷液后回吸体积指令
            if (cmd.ToLower().Contains(ConstantsUtils.BackAbsorbCmd.ToLower()))
                return true;

            // 剔除吸液后多吸体积指令
            if (cmd.ToLower().Contains(ConstantsUtils.AbsorbMoreCmd.ToLower()))
                return true;

            // 剔除吸后反喷体积指令
            if (cmd.ToLower().Contains(ConstantsUtils.ReverseJetCmd.ToLower()))
                return true;

            // 剔除多吸液体返回源孔喷出指令
            if (cmd.ToLower().Contains(ConstantsUtils.ReJet2SourceCmd.ToLower()))
                return true;

            return false;
        }

        /// <summary>
        /// 从Enum中获取描述Description
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string GetDescription(this System.Enum value)
        {
            var fieldInfo = value.GetType().GetField(value.ToString());
            var attributes = (DescriptionAttribute[])fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);

            return attributes.Length > 0 ? attributes[0].Description : value.ToString();
        }

        /// <summary>
        /// 获取运动轴相应指令（即如果包含@，就转换）
        /// </summary>
        /// <param name=""></param>
        /// <param name=""></param>
        /// <param name=""></param>
        /// <param name=""></param>
        /// <param name=""></param>
        /// <param name=""></param>
        /// <returns></returns>
        public static string GetMotionCmd(string axis, EActType actType, string value)
        {
            var cmd = "";
            if (axis.Contains("@")) cmd = "X" + actType.GetDescription() + value + axis;
            else cmd = axis + actType.GetDescription() + value;
            return cmd;
        }
}
}
