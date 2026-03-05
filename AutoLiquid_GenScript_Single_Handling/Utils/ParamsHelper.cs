using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xaml.Permissions;
using AutoLiquid_GenScript_Single_Handling.EntityJson;
using AutoLiquid_Library.Enum;
using Newtonsoft.Json;

namespace AutoLiquid_GenScript_Single_Handling.Utils
{
    /// <summary>
    /// 参数工具类
    /// </summary>
    public class ParamsHelper
    {
        // 默认耗材类型数量
        private static int defaultConsumableCount = 5;

        // 盘位布局
        public static Layout Layout;
        // 耗材设置
        public static List<Common> CommonSettingList = new List<Common> { new Common(), new Common() };
        // 移液头
        public static List<Head> HeadList = new List<Head> { new Head(), new Head() };
        // 量程最大值
        public static Range Range;
        // 输入输出
        public static IO IO;
        // 标准原点
        public static OriginalPoint OriginalPoint;
        // 机器偏移值
        public static Offsets Offsets;
        // 权限管理
        public static Permission Permission;
        // 调试参数
        public static Debug Debug;

        /// <summary>
        /// 加载所有参数
        /// </summary>
        public static bool LoadAllParams()
        {
            try
            {
                LoadLayout();
                LoadCommonSetting(0);
                LoadCommonSetting(1);
                LoadHead(0);
                LoadHead(1);
                LoadRange();
                LoadIO();
                LoadOriginalPoint();
                LoadOffsets();
                LoadPermission();
                LoadDebug();
            }
            catch (Exception e)
            {
                LogHelper.Error("解析参数失败：" + e.Message + Environment.NewLine + e.StackTrace);
                return false;
            }
            return true;
        }

        private static void LoadLayout()
        {
            String jsonStr =
                FileUtils.GetJsonFromLocalStorage(AutoLiquid_Library.Utils.ConstantsUtils.FILE_LAYOUT);
            if (jsonStr.Trim() == "") // 未有参数文件
            {
                Layout = new Layout();
                FileUtils.SaveLayout(Layout);
            }
            else
            {
                Layout = JsonConvert.DeserializeObject<Layout>(jsonStr);
            }
        }

        /// <summary>
        /// 移液头耗材设置
        /// </summary>
        /// <param name="headIndex"></param>
        private static void LoadCommonSetting(int headIndex)
        {
            // 文件名
            var jsonName = headIndex == 0
                ? AutoLiquid_Library.Utils.ConstantsUtils.FILE_COMMON_SETTING
                : AutoLiquid_Library.Utils.ConstantsUtils.FILE_COMMON_SETTING_2;
            Common commonSetting;

            String jsonStr = FileUtils.GetJsonFromLocalStorage(jsonName);
            if (jsonStr.Trim() == "") // 未有参数文件
            {
                commonSetting = new Common();
                AddConsumable(commonSetting);
                AddReleaseTipPos(commonSetting);
                FileUtils.SaveCommonSettings(headIndex, commonSetting);
            }
            else
            {
                commonSetting = JsonConvert.DeserializeObject<Common>(jsonStr);

                /**
                 * 以下代码为了兼容盘位不足等问题
                 */
                var templateCount = ParamsHelper.Layout.RowCount * ParamsHelper.Layout.ColCount;
                // 预防耗材为空
                if (commonSetting.Consumables.Count == 0)
                {
                    AddConsumable(commonSetting);
                    FileUtils.SaveCommonSettings(headIndex, commonSetting);
                }
                // 预防耗材HoleStartPosList不足（以第1个耗材为准）
                else if (commonSetting.Consumables[0].HoleStartPosList.Count < templateCount)
                {
                    // 缺少的盘位数
                    var templateCountLack = templateCount - commonSetting.Consumables[0].HoleStartPosList.Count;
                    foreach (var commonGroup in commonSetting.Consumables)
                    {
                        AddPositionHole(commonGroup.HoleStartPosList, templateCountLack);
                    }
                    FileUtils.SaveCommonSettings(headIndex, commonSetting);
                }

                // 预防耗材盘位使能属性为空（以第1个耗材为准）
                if (commonSetting.Consumables[0].TemplateAvailableList.Count == 0)
                {
                    AddTemplateAvailable(commonSetting, templateCount);
                    FileUtils.SaveCommonSettings(headIndex, commonSetting);
                }
                // 预防耗材TemplateAvailableList不足（以第1个耗材为准）
                else if (commonSetting.Consumables[0].TemplateAvailableList.Count < templateCount)
                {
                    // 缺少的盘位数
                    var templateCountLack = templateCount - commonSetting.Consumables[0].TemplateAvailableList.Count;
                    AddTemplateAvailable(commonSetting, templateCountLack);
                    FileUtils.SaveCommonSettings(headIndex, commonSetting);
                }

                // 预防第1个耗材不是枪头盒
                if (!commonSetting.Consumables[0].IsTipBox)
                    commonSetting.Consumables[0].IsTipBox = true;

                // 预防退枪头位置为空
                if (commonSetting.ReleaseTipPosList.Count == 0)
                    AddReleaseTipPos(commonSetting);
            }

            CommonSettingList[headIndex] = commonSetting;
        }

        /// <summary>
        /// 添加耗材
        /// </summary>
        /// <param name="commonSetting"></param>
        private static void AddConsumable(Common commonSetting)
        {
            for (var i = 0; i < defaultConsumableCount; i++)
            {
                Consumable cg = new Consumable();
                // 第一组默认为枪头盒
                if (i == 0)
                {
                    cg.GroupName = (string)Application.Current.FindResource("TemplateTips");
                    cg.IsTipBox = true;
                }
                else
                    cg.GroupName = (string)Application.Current.FindResource("Consumable") + (i + 1);
                commonSetting.Consumables.Add(cg);

                var templateCount = ParamsHelper.Layout.RowCount * ParamsHelper.Layout.ColCount;
                AddPositionHole(commonSetting.Consumables[i].HoleStartPosList, templateCount);
                AddTemplateAvailableSub(commonSetting.Consumables[i].TemplateAvailableList, templateCount);
            }
        }

        /// <summary>
        /// 添加退枪头位置
        /// </summary>
        /// <param name="commonSetting"></param>
        public static void AddReleaseTipPos(Common commonSetting)
        {
            AddPositionHole(commonSetting.PrepareReleaseTipPosList, 4);
            AddPositionHole(commonSetting.ReleaseTipPosList, 4);
            for (var i = 0; i < 4; i++)
            {
                commonSetting.ReleaseTipPosAvailableList.Add(i == 0);
            }
        }

        /// <summary>
        /// 添加孔位
        /// </summary>
        /// <param name="posList"></param>
        /// <param name="posCount"></param>
        public static void AddPositionHole(List<Position> posList, int posCount)
        {
            for (var i = 0; i < posCount; i++)
            {
                var pos = new Position { X = 0.00m, Y = 0.00m };
                posList.Add(pos);
            }
        }

        /// <summary>
        /// 添加耗材盘位使能
        /// </summary>
        /// <param name="commonSetting"></param>
        /// <param name="addTemplateCount">需要添加的盘位数目</param>
        private static void AddTemplateAvailable(Common commonSetting, int addTemplateCount)
        {
            foreach (var commonGroup in commonSetting.Consumables)
            {
                AddTemplateAvailableSub(commonGroup.TemplateAvailableList, addTemplateCount);
            }
        }

        public static void AddTemplateAvailableSub(List<bool> availableList, int templateCount)
        {
            for (var i = 0; i < templateCount; i++)
            {
                availableList.Add(true);
            }
        }


        /// <summary>
        /// 移液头设置
        /// </summary>
        private static void LoadHead(int headIndex)
        {
            // 文件名
            var jsonName = headIndex == 0
                ? AutoLiquid_Library.Utils.ConstantsUtils.FILE_HEAD_1
                : AutoLiquid_Library.Utils.ConstantsUtils.FILE_HEAD_2;
            Head head;

            String jsonStr =
                FileUtils.GetJsonFromLocalStorage(jsonName);
            if (jsonStr.Trim() == "") // 未有参数文件
            {
                head = new Head();
                // 移液头2默认不可用
                if (headIndex == 1)
                    head.Available = false;
                FileUtils.SaveHead(headIndex, head);
            }
            else
            {
                head = JsonConvert.DeserializeObject<Head>(jsonStr);
                // 主动生成没有的参数
                if (!jsonStr.Contains("ReleaseTipUsePushCount"))
                {
                    FileUtils.SaveHead(headIndex, head);
                }
            }

            HeadList[headIndex] = head;
        }

        /// <summary>
        /// 最大量程设置
        /// </summary>
        private static void LoadRange()
        {
            String jsonStr =
                FileUtils.GetJsonFromLocalStorage(AutoLiquid_Library.Utils.ConstantsUtils.FILE_RANGE);
            if (jsonStr.Trim() == "") // 未有参数文件
            {
                Range = new Range();
                FileUtils.SaveRange(Range);
            }
            else
            {
                Range = JsonConvert.DeserializeObject<Range>(jsonStr);
            }
        }

        /// <summary>
        /// 输入输出设置
        /// </summary>
        private static void LoadIO()
        {
            String jsonStr =
                FileUtils.GetJsonFromLocalStorage(AutoLiquid_Library.Utils.ConstantsUtils.FILE_IO);
            if (jsonStr.Trim() == "") // 未有参数文件
            {
                IO = new IO();
                FileUtils.SaveIO(IO);
            }
            else
            {
                IO = JsonConvert.DeserializeObject<IO>(jsonStr);
            }
        }

        /// <summary>
        /// 标准原点
        /// </summary>
        private static void LoadOriginalPoint()
        {
            String jsonStr =
                FileUtils.GetJsonFromLocalStorage(AutoLiquid_Library.Utils.ConstantsUtils.FILE_ORIGINAL_POINT);
            if (jsonStr.Trim() == "") // 未有参数文件
            {
                OriginalPoint = new OriginalPoint();
                FileUtils.SaveOriginalPoint(OriginalPoint);
            }
            else
            {
                OriginalPoint = JsonConvert.DeserializeObject<OriginalPoint>(jsonStr);
            }
        }

        /// <summary>
        /// 偏移值设置
        /// </summary>
        private static void LoadOffsets()
        {
            String jsonStr =
                FileUtils.GetJsonFromLocalStorage(AutoLiquid_Library.Utils.ConstantsUtils.FILE_OFFSETS);
            if (jsonStr.Trim() == "") // 未有参数文件
            {
                Offsets = new Offsets();
                FileUtils.SaveOffsets(Offsets);
            }
            else
            {
                Offsets = JsonConvert.DeserializeObject<Offsets>(jsonStr);
            }
        }

        /// <summary>
        /// 权限管理
        /// </summary>
        private static void LoadPermission()
        {
            String jsonStr =
                FileUtils.GetJsonFromLocalStorage(AutoLiquid_Library.Utils.ConstantsUtils.FILE_PERMISSION);
            if (jsonStr.Trim() == "") // 未有参数文件
            {
                // 创建默认密码：123321
                Permission = new Permission { PwdHash = PermissionHelper.CreateHash("123321") };
                FileUtils.SavePermission(Permission);
            }
            else
            {
                Permission = JsonConvert.DeserializeObject<Permission>(jsonStr);
            }
        }

        /// <summary>
        /// 调试参数
        /// </summary>
        private static void LoadDebug()
        {
            String jsonStr =
                FileUtils.GetJsonFromLocalStorage(AutoLiquid_Library.Utils.ConstantsUtils.FILE_DEBUG);
            if (jsonStr.Trim() == "") // 未有参数文件
            {
                Debug = new Debug();
                FileUtils.SaveDebug(Debug);
            }
            else
            {
                Debug = JsonConvert.DeserializeObject<Debug>(jsonStr);
            }
        }

    }
}
