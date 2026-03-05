using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using AutoLiquid_Library.Enum;
using DAERun;

namespace AutoLiquid_GenScript_Single_Handling.Utils
{
    /// <summary>
    /// 数据处理帮助类
    /// </summary>
    public class DataHelper
    {
        /// <summary>
        /// 保存整数值
        /// </summary>
        /// <param name="headIndex">移液头Index</param>
        /// <param name="inputValue">保存的值</param>
        /// <param name="param">Json属性</param>
        /// <param name="saveAction">保存方法</param>
        public static void SaveInt(int headIndex, string inputValue, ref int param, Action saveAction = null)
        {
            // 避免输入负号出现异常
            if (inputValue.Equals("") || inputValue.Equals("-"))
                return;

            var result = int.TryParse(inputValue, out var outputValue);
            if (result)
            {
                param = outputValue;
                if (saveAction == null)
                    FileUtils.SaveCommonSettings(headIndex, ParamsHelper.CommonSettingList[headIndex]);
                else
                    saveAction.Invoke();
            }
            else
            {
                // MessageBox要在主线程运行，否则报错
                Application.Current.Dispatcher.Invoke(() =>
                {
                    MessageBox.Show((string)Application.Current.FindResource("Prompt_Data_Input_Error"));
                });
            }
        }

        /// <summary>
        /// 保存小数值
        /// </summary>
        /// <param name="headIndex">移液头Index</param>
        /// <param name="inputValue">保存的值</param>
        /// <param name="param">Json属性</param>
        /// <param name="saveAction">保存方法</param>
        public static void SaveDecimal(int headIndex, string inputValue, ref decimal param, Action saveAction = null)
        {
            // 避免输入负号出现异常
            if(inputValue.Equals("") || inputValue.Equals("-"))
                return;

            var result = decimal.TryParse(inputValue, out var outputValue);
            if (result)
            {
                param = outputValue;
                if (saveAction == null)
                    FileUtils.SaveCommonSettings(headIndex, ParamsHelper.CommonSettingList[headIndex]);
                else
                    saveAction.Invoke();
            }
            else
            {
                // MessageBox要在主线程运行，否则报错
                Application.Current.Dispatcher.Invoke(() =>
                {
                    MessageBox.Show((string)Application.Current.FindResource("Prompt_Data_Input_Error"));
                });
            }
        }


        /// <summary>
        /// 保存字符串
        /// </summary>
        /// <param name="headIndex">移液头Index</param>
        /// <param name="inputValue">保存的值</param>
        /// <param name="param">Json属性</param>
        /// <param name="saveAction">保存方法</param>
        public static void SaveString(int headIndex, string inputValue, ref string param, Action saveAction = null)
        {
            param = inputValue;
            if (saveAction == null)
                FileUtils.SaveCommonSettings(headIndex, ParamsHelper.CommonSettingList[headIndex]);
            else
                saveAction.Invoke();
        }

        /// <summary>
        /// 保存布尔值
        /// </summary>
        /// <param name="headIndex">移液头Index</param>
        /// <param name="inputValue">保存的值</param>
        /// <param name="param">Json属性</param>
        /// <param name="saveAction">保存方法</param>
        public static void SaveBool(int headIndex, bool inputValue, ref bool param, Action saveAction = null)
        {
            param = inputValue;
            if (saveAction == null)
                FileUtils.SaveCommonSettings(headIndex, ParamsHelper.CommonSettingList[headIndex]);
            else
                saveAction.Invoke();
        }

        /// <summary>
        /// 保存量程值
        /// </summary>
        /// <param name="inputValue">保存的值</param>
        /// <param name="param">Json属性</param>
        /// <param name="saveAction">保存方法</param>
        public static void SaveLiquidRange(ELiquidRange inputValue, ref ELiquidRange param, Action saveAction)
        {
            param = inputValue;
            saveAction.Invoke();
        }

        /// <summary>
        /// 保存盘位摆放方向
        /// </summary>
        /// <param name="inputValue">保存的值</param>
        public static void SaveA1Pos(EA1Pos inputValue)
        {
            ParamsHelper.CommonSettingList[0].A1Pos = inputValue;
            FileUtils.SaveCommonSettings(0, ParamsHelper.CommonSettingList[0]);
            ParamsHelper.CommonSettingList[1].A1Pos = inputValue;
            FileUtils.SaveCommonSettings(1, ParamsHelper.CommonSettingList[1]);
        }

        /// <summary>
        /// 对体积vol进行校准
        /// </summary>
        /// <param name="headUsedIndex"></param>
        /// <param name="vol"></param>
        /// <returns></returns>
        public static decimal CalibrateVol(int headUsedIndex, decimal vol)
        {
            var multiCalibration = ParamsHelper.CommonSettingList[headUsedIndex].MultiCalibration;

            if (!multiCalibration.Available)
                return vol;

            // 校准体积
            var pVolArray = multiCalibration.PVolArray;
            // 补偿体积
            var pCompensationArray = multiCalibration.PCompensationArray;

            // 对应最小的补偿值
            decimal vCompensation = pCompensationArray.Last();

            if (vol >= pVolArray[0])
            {
                vCompensation = pCompensationArray[0];
            }
            else if (vol >= pVolArray[1] && vol < pVolArray[0])
            {
                vCompensation = pCompensationArray[1] + (vol - pVolArray[1]) * (pCompensationArray[0] - pCompensationArray[1]) / (pVolArray[0] - pVolArray[1]);
            }
            else if (vol >= pVolArray[2] && vol < pVolArray[1])
            {
                vCompensation = pCompensationArray[2] + (vol - pVolArray[2]) * (pCompensationArray[1] - pCompensationArray[2]) / (pVolArray[1] - pVolArray[2]);
            }
            else if (vol >= pVolArray[3] && vol < pVolArray[2])
            {
                vCompensation = pCompensationArray[3] + (vol - pVolArray[3]) * (pCompensationArray[2] - pCompensationArray[3]) / (pVolArray[2] - pVolArray[3]);
            }
            else if (vol >= pVolArray[4] && vol < pVolArray[3])
            {
                vCompensation = pCompensationArray[4] + (vol - pVolArray[4]) * (pCompensationArray[3] - pCompensationArray[4]) / (pVolArray[3] - pVolArray[4]);
            }

            var result = Math.Round(vol + vCompensation, 2);

            return result;
        }

        /// <summary>
        /// 获取字符串中的时分秒，并转换为总秒数
        /// </summary>
        /// <param name="timeString"></param>
        /// <returns></returns>
        public static int GetTotalSeconds(string timeString)
        {
            int totalSeconds = 0;
            int hours = 0;
            int minutes = 0;
            int seconds = 0;
            int index = 0;

            timeString = timeString.ToLower().Replace("wait", "").Trim();
            // 解析小时  
            index = timeString.IndexOf('h');
            // 确保 'h' 不是字符串的第一个字符  
            if (index > -1 && index > 0) 
            {
                if (int.TryParse(timeString.Substring(0, index), out hours))
                {
                    // 转换为秒  
                    totalSeconds += hours * 3600;
                    // 去掉已解析的部分和可能的空格
                    timeString = timeString.Substring(index + 1).TrimStart();
                }
            }

            // 解析分钟  
            index = timeString.IndexOf('m');
            // 确保 'm' 不是字符串的第一个字符  
            if (index > -1 && index > 0) 
            {
                if (int.TryParse(timeString.Substring(0, index), out minutes))
                {
                    // 转换为秒  
                    totalSeconds += minutes * 60;
                    // 去掉已解析的部分和可能的空格  
                    timeString = timeString.Substring(index + 1).TrimStart();
                }
            }

            // 解析秒  
            index = timeString.IndexOf('s');
            // 确保 's' 不是字符串的第一个字符  
            if (index > -1 && index > 0)
            {
                if (int.TryParse(timeString.Substring(0, index), out seconds))
                {
                    // 转换为秒  
                    totalSeconds += seconds;
                }
            }

            return totalSeconds;
        }

        /// <summary>
        /// List拆分成多个子List，每个子List对应的Index数值之和等于List对应Index值
        /// 例如List = {10m, 20m, 30m, 5m}, 阈值为10，则拆分成多个subList = {10, 10, 10, 5},{0, 10, 10, 0},{0, 0, 10, 0}
        /// </summary>
        /// <param name="list"></param>
        /// <param name="threshold"></param>
        /// <returns></returns>
        public static List<List<decimal>> SplitList(List<decimal> list, decimal threshold)
        {
            List<List<decimal>> sublists = new List<List<decimal>>();
            // 需要分拆次数
            int maxSplits = 0;

            // 计算分拆次数
            foreach (var item in list)
            {
                // 使用 Math.Ceiling 来进行向上取整
                maxSplits = Math.Max(maxSplits, (int)Math.Ceiling(item / threshold));
            }

            for (int splitIndex = 0; splitIndex < maxSplits; splitIndex++)
            {
                List<decimal> sublist = new List<decimal>();
                for (int i = 0; i < list.Count; i++)
                {
                    decimal remainingValue = list[i];
                    if (remainingValue >= threshold)
                    {
                        list[i] -= threshold;
                        sublist.Add(threshold);
                    }
                    else
                    {
                        sublist.Add(remainingValue);
                        list[i] = 0m;
                    }
                }
                sublists.Add(sublist);
            }

            return sublists;
        }
    }
}
