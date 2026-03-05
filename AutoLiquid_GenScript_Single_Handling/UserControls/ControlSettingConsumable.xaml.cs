using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using AutoLiquid_GenScript_Single_Handling.EntityJson;
using AutoLiquid_GenScript_Single_Handling.Utils;
using AutoLiquid_GenScript_Single_Handling.Window;
using AutoLiquid_Library.Utils;
using MahApps.Metro.Controls;

namespace AutoLiquid_GenScript_Single_Handling.UserControls
{
    /// <summary>
    /// 耗材定义控件
    /// </summary>
    public partial class ControlSettingConsumable : UserControl
    {
        // 移液头Index
        private int mHeadIndex;

        // 耗材详细界面
        private List<ControlSettingGroup> ControlSettingGroups = new List<ControlSettingGroup>();

        // 上一个选择的标签Index
        private int lastTabSelectedIndex = 0;

        public ControlSettingConsumable(int headIndex)
        {
            InitializeComponent();

            this.mHeadIndex = headIndex;

            this.Loaded += OnLoaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            // 初始化控件
            InitWidget();
            // 控件事件
            ControlEvent();
        }

        private void InitWidget()
        {
            for (var i = 0; i < ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables.Count; i++)
            {
                AddTabItem(i, ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables[i].IsTipBox);
            }
            // 标签添加按钮（最后一个标签）
            this.TabControlGroup.Items.Add(new MetroTabItem { Header = "+", Foreground = Brushes.LightSkyBlue, Background = Brushes.Transparent });
            // 刷新标签名
            RefreshGroupTitle();
        }

        private void ControlEvent()
        {
            // 标签相关事件
            ControlTabEvent();
        }

        /// <summary>
        /// 添加TabItem
        /// </summary>
        /// <param name="tabIndex"></param>
        /// <param name="isTipBox">是否枪头盒</param>
        private void AddTabItem(int tabIndex, bool isTipBox)
        {
            ControlSettingGroups.Add(new ControlSettingGroup(tabIndex, this.mHeadIndex, isTipBox));
            var metroTabStyle = Application.Current.FindResource("MetroTabStyle") as Style;
            var metroTabStyleWithClose = Application.Current.FindResource("MetroTabStyleWithClose") as Style;
            var tabItem = tabIndex == 0 ? new MetroTabItem { Style = metroTabStyle } : new MetroTabItem { Style = metroTabStyleWithClose };
            this.TabControlGroup.Items.Insert(tabIndex, tabItem);
            // 显示页面
            ((MetroTabItem)this.TabControlGroup.Items[tabIndex]).Content = ControlSettingGroups.ElementAt(tabIndex);
        }

        /// <summary>
        /// 插入TabItem
        /// </summary>
        /// <param name="tabIndex"></param>
        /// <param name="isTipBox">是否枪头盒</param>
        private void InsertTabItem(int tabIndex, bool isTipBox)
        {
            ControlSettingGroups.Insert(tabIndex, new ControlSettingGroup(tabIndex, this.mHeadIndex, isTipBox));
            var metroTabStyle = Application.Current.FindResource("MetroTabStyle") as Style;
            var metroTabStyleWithClose = Application.Current.FindResource("MetroTabStyleWithClose") as Style;
            var tabItem = new MetroTabItem { Style = metroTabStyleWithClose };
            this.TabControlGroup.Items.Insert(tabIndex, tabItem);
            // 显示页面
            ((MetroTabItem)this.TabControlGroup.Items[tabIndex]).Content = ControlSettingGroups.ElementAt(tabIndex);
        }

        /// <summary>
        /// 刷新分组标题
        /// </summary>
        private void RefreshGroupTitle()
        {
            for (var i = 0; i < ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables.Count; i++)
            {
                var commonGroup = ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables[i];
                ((MetroTabItem)this.TabControlGroup.Items[i]).Header = commonGroup.GroupName;
            }
        }

        private void ControlTabEvent()
        {
            // 耗材标签双击事件
            for (var i = 0; i < ControlSettingGroups.Count; i++)
            {
                ControlMouseDoubleClickEvent(i);
            }

            // 标签关闭事件
            this.TabControlGroup.TabItemClosingEvent += TabControlOnTabItemClosingEvent;

            // 标签添加点击事件
            this.TabControlGroup.SelectionChanged += TabControlGroupOnSelectionChanged;
        }

        private void ControlMouseDoubleClickEvent(int tabIndex)
        {
            // 第一个枪头盒双击添加枪头盒耗材
            if (tabIndex == 0)
                ((MetroTabItem)this.TabControlGroup.Items[tabIndex]).MouseDoubleClick += TabItemGroupAddTipBoxOnMouseDoubleClick;
            // 其他耗材双击修改名称
            else
                ((MetroTabItem)this.TabControlGroup.Items[tabIndex]).MouseDoubleClick += TabItemGroupNameModifyOnMouseDoubleClick;
        }

        private void TabControlOnTabItemClosingEvent(object sender, BaseMetroTabControl.TabItemClosingEventArgs e)
        {
            var closingTabIndex = ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables.FindIndex(p => p.GroupName.Equals(e.ClosingTabItem.Header.ToString()));
            if (closingTabIndex < 0)
                return;
            if (MessageBox.Show((string)this.FindResource("Prompt_If_Delete_Consumable_1") + ((TabItem)this.TabControlGroup.Items[closingTabIndex]).Header + (string)this.FindResource("Prompt_If_Delete_Consumable_2"),
                    (string)this.FindResource("Prompt"), MessageBoxButton.YesNo, MessageBoxImage.Warning) ==
                MessageBoxResult.No)
            {
                e.Cancel = true;
            }
            else
            {
                ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables.RemoveAt(closingTabIndex);
                FileUtils.SaveCommonSettings(this.mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex]);
                ControlSettingGroups.RemoveAt(closingTabIndex);

                // 避免已选的Tab和关闭的Tab相同时，标签删除不了
                if (this.TabControlGroup.SelectedIndex == closingTabIndex)
                    this.TabControlGroup.SelectedIndex = 0;
            }
        }

        private void TabControlGroupOnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 如果点击“+”标签，就选中倒数第二个Tab
            if (this.TabControlGroup.SelectedIndex == this.TabControlGroup.Items.Count - 1)
            {
                var checkGroupNameTick = 0;
                var groupNameNew = "";
                while (true)
                {
                    checkGroupNameTick++;
                    groupNameNew = (string)this.FindResource("Consumable") +
                                      (ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables.Count + checkGroupNameTick);
                    if (!ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables.Exists(p => p.GroupName.Equals(groupNameNew)))
                        break;
                }

                // 新耗材默认参数复制当前选中的tab，否则复制最后一个tab
                var newCommonGroup = new Consumable();
                var commonGroupsCount = ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables.Count;
                if (commonGroupsCount > 1)
                {
                    if (lastTabSelectedIndex != this.TabControlGroup.Items.Count - 1 && !ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables[lastTabSelectedIndex].IsTipBox)
                        newCommonGroup = ObjectUtils.DeepCopy(ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables[lastTabSelectedIndex]);
                    else
                        newCommonGroup = ObjectUtils.DeepCopy(ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables[commonGroupsCount - 1]);
                }
                newCommonGroup.GroupName = groupNameNew;
                ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables.Add(newCommonGroup);

                var templateCount = ParamsHelper.Layout.RowCount * ParamsHelper.Layout.ColCount;
                var groupIndex = ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables.Count - 1;
                ParamsHelper.AddPositionHole(ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables[groupIndex].HoleStartPosList, templateCount);
                ParamsHelper.AddTemplateAvailableSub(
                    ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables[groupIndex].TemplateAvailableList, templateCount);
                FileUtils.SaveCommonSettings(this.mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex]);

                // 默认插入到倒数第二个Tab
                AddTabItem(this.TabControlGroup.Items.Count - 1, false);
                var tabSelectedIndex = this.TabControlGroup.Items.Count - 2;
                this.TabControlGroup.SelectedIndex = tabSelectedIndex;

                RefreshGroupTitle();
                ControlMouseDoubleClickEvent(tabSelectedIndex);
            }

            lastTabSelectedIndex = this.TabControlGroup.SelectedIndex;
        }

        /// <summary>
        /// 添加枪头盒耗材
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TabItemGroupAddTipBoxOnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Point position = e.GetPosition(this);
            double mouseY = position.Y;

            // y方向小于等于40像素点，才执行动作，避免误双击
            if (mouseY <= 40)
            {
                var groupIndexNew = ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables.Count(p => p.IsTipBox);
                var groupNameNew = (string)this.FindResource("TemplateTips") + (groupIndexNew + 1);
                // 新枪头盒默认参数跟第1个枪头盒耗材一样
                var newCommonGroup =
                    ObjectUtils.DeepCopy(ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables[0]);
                newCommonGroup.GroupName = groupNameNew;
                ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables.Insert(groupIndexNew, newCommonGroup);
                var templateCount = ParamsHelper.Layout.RowCount * ParamsHelper.Layout.ColCount;
                ParamsHelper.AddPositionHole(ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables[groupIndexNew].HoleStartPosList, templateCount);
                ParamsHelper.AddTemplateAvailableSub(
                    ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables[groupIndexNew].TemplateAvailableList, templateCount);
                FileUtils.SaveCommonSettings(this.mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex]);

                // 插入TabItem
                InsertTabItem(groupIndexNew, true);
                this.TabControlGroup.SelectedIndex = groupIndexNew;

                RefreshGroupTitle();
                ControlMouseDoubleClickEvent(groupIndexNew);
            }
        }

        /// <summary>
        /// 修改耗材名称
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TabItemGroupNameModifyOnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Point position = e.GetPosition(this);
            double mouseY = position.Y;

            // y方向小于等于40像素点，才执行动作，避免误双击
            if (mouseY <= 40)
            {
                var doubleClickTabIndex = ParamsHelper.CommonSettingList[this.mHeadIndex].Consumables.FindIndex(p => p.GroupName.Equals(((MetroTabItem)sender).Header.ToString()));
                if (doubleClickTabIndex < 0)
                    return;
                if (sender.Equals(((MetroTabItem)this.TabControlGroup.Items[doubleClickTabIndex])))
                {
                    if (WindowChangeGroupTitle.EnsureExecute(this.mHeadIndex, doubleClickTabIndex))
                        RefreshGroupTitle();
                }
            }
        }
    }
}
