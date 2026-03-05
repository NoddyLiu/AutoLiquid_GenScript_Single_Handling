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
using AutoLiquid_GenScript_Single_Handling.Utils;
using AutoLiquid_Library.Enum;
using DataHelper = AutoLiquid_GenScript_Single_Handling.Utils.DataHelper;

namespace AutoLiquid_GenScript_Single_Handling.UserControls
{
    /// <summary>
    /// 退枪头控件
    /// </summary>
    public partial class ControlSettingReleaseTip : UserControl
    {
        // 移液头Index
        private int mHeadIndex;

        // 退枪头位置List
        private List<ControlSettingReleaseTipPos> ControlSettingReleaseTipPosList = new List<ControlSettingReleaseTipPos>();

        public ControlSettingReleaseTip(int headIndex)
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
            this.RBtnReleaseTipMethodPush.IsChecked = ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipUsePush;
            this.RBtnReleaseTipMethodHook.IsChecked = !ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipUsePush;
            this.RBtnReleaseTipAxisP.IsChecked = ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipAxis == EAxis.P;
            this.RBtnReleaseTipAxisQ.IsChecked = ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipAxis == EAxis.Q;
            this.ComboBoxPushCount.SelectedItem = ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipUsePushCount;
            this.ControlReleaseTipOffset.SetParam(mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipOffset, ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipAxis);
            this.TextBoxReleaseTipVariableDistanceStep.Text = ParamsHelper.CommonSettingList[this.mHeadIndex]
                .ReleaseTipVariableDistanceStep.ToString();

            // 退枪头位置
            ControlSettingReleaseTipPosList.AddRange(new List<ControlSettingReleaseTipPos> { this.ControlSettingReleaseTipPos1, this.ControlSettingReleaseTipPos2, this.ControlSettingReleaseTipPos3, this.ControlSettingReleaseTipPos4 });
            for (var i = 0; i < ControlSettingReleaseTipPosList.Count; i++)
            {
                if (i == 0)
                    ControlSettingReleaseTipPosList[i].CheckBoxReleaseTipPos.Visibility = Visibility.Collapsed;

                ControlSettingReleaseTipPosList[i].LabelPos.Content = (i + 1).ToString();
                ControlSettingReleaseTipPosList[i].ControlPrepareReleaseTipPosX.SetParam(mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[i].X, EAxis.X);
                ControlSettingReleaseTipPosList[i].ControlPrepareReleaseTipPosY.SetParam(mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[i].Y, EAxis.Y);
                ControlSettingReleaseTipPosList[i].ControlPrepareReleaseTipPosZ.SetParam(mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[i].Z, EAxis.Z);
                ControlSettingReleaseTipPosList[i].ControlReleaseTipPosX.SetParam(mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[i].X, EAxis.X);
                ControlSettingReleaseTipPosList[i].ControlReleaseTipPosY.SetParam(mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[i].Y, EAxis.Y);
                ControlSettingReleaseTipPosList[i].ControlReleaseTipPosZ.SetParam(mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[i].Z, EAxis.Z);

                // 退枪头位置可用状态
                RefreshReleaseTipPosAvailable(i);
            }

            this.CheckBoxAxisXGoFirst.IsChecked = ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipAxisXGoFirst;
            this.CheckBoxAxisYGoFirst.IsChecked = ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipAxisYGoFirst;
            this.CheckBoxReleaseTipBack2TakePos.IsChecked = ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipBack2TakePos;
            this.CheckBoxReleaseTipZa0Before.IsChecked = ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipZa0Before;
            this.CheckBoxReleaseTipZa0After.IsChecked = ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipZa0After;
            this.TextBoxReleaseTipSpeedCmd.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipSpeedCmd;
            this.TextBoxReleaseTipAfterCmd.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipAfterCmd;

            // 如果移液头不是推脱板退枪头，隐藏相关属性
            this.StackPanelReleaseTipPush.Visibility = ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipUsePush ? Visibility.Visible : Visibility.Collapsed;

            // 如果移液头2可用，隐藏推脱板轴属性
            this.StackPanelReleaseTipAxis.Visibility = ParamsHelper.HeadList[1].Available ? Visibility.Collapsed : Visibility.Visible;

            // 如果移液头可变距，且退枪头方式为卡扣，显示退枪头变距属性
            this.StackPanelReleaseTipVariableDistanceStep.Visibility =
                ParamsHelper.HeadList[this.mHeadIndex].IsVariable &&
                !ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipUsePush
                    ? Visibility.Visible
                    : Visibility.Collapsed;
        }

        /// <summary>
        /// 刷新退枪头位置可用状态
        /// </summary>
        /// <param name="posIndex"></param>
        private void RefreshReleaseTipPosAvailable(int posIndex)
        {
            if (ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosAvailableList[posIndex])
            {
                ControlSettingReleaseTipPosList[posIndex].CheckBoxReleaseTipPos.IsChecked = true;
                ViewUtils.SetEnableExceptCheckbox(ControlSettingReleaseTipPosList[posIndex].GridReleaseTipPos, true);
            }
            else
            {
                ControlSettingReleaseTipPosList[posIndex].CheckBoxReleaseTipPos.IsChecked = false;
                ViewUtils.SetEnableExceptCheckbox(ControlSettingReleaseTipPosList[posIndex].GridReleaseTipPos, false);
            }
        }

        private void ControlEvent()
        {
            this.RBtnReleaseTipMethodPush.Checked += RBtnOnChecked;
            this.RBtnReleaseTipMethodHook.Checked += RBtnOnChecked;
            this.RBtnReleaseTipAxisP.Checked += RBtnOnChecked;
            this.RBtnReleaseTipAxisQ.Checked += RBtnOnChecked;
            this.ComboBoxPushCount.SelectionChanged += ComboBoxOnSelectionChanged;
            this.ControlReleaseTipOffset.TextBoxPos.TextChanged += TextBoxOnTextChanged;
            this.TextBoxReleaseTipVariableDistanceStep.TextChanged += TextBoxOnTextChanged;
            this.BtnReleaseTipVariableDistanceStep.Click += BtnOnClick;

            for (var i = 0; i < ControlSettingReleaseTipPosList.Count; i++)
            {
                ControlSettingReleaseTipPosList[i].ControlPrepareReleaseTipPosX.TextBoxPos.TextChanged += TextBoxOnTextChanged;
                ControlSettingReleaseTipPosList[i].ControlPrepareReleaseTipPosY.TextBoxPos.TextChanged += TextBoxOnTextChanged;
                ControlSettingReleaseTipPosList[i].ControlPrepareReleaseTipPosZ.TextBoxPos.TextChanged += TextBoxOnTextChanged;
                ControlSettingReleaseTipPosList[i].ControlReleaseTipPosX.TextBoxPos.TextChanged += TextBoxOnTextChanged;
                ControlSettingReleaseTipPosList[i].ControlReleaseTipPosY.TextBoxPos.TextChanged += TextBoxOnTextChanged;
                ControlSettingReleaseTipPosList[i].ControlReleaseTipPosZ.TextBoxPos.TextChanged += TextBoxOnTextChanged;
                ControlSettingReleaseTipPosList[i].CheckBoxReleaseTipPos.Checked += CheckBoxOnChecked;
                ControlSettingReleaseTipPosList[i].CheckBoxReleaseTipPos.Unchecked += CheckBoxOnUnChecked;
            }

            this.CheckBoxAxisXGoFirst.Checked += CheckBoxOnChecked;
            this.CheckBoxAxisXGoFirst.Unchecked += CheckBoxOnUnChecked;
            this.CheckBoxAxisYGoFirst.Checked += CheckBoxOnChecked;
            this.CheckBoxAxisYGoFirst.Unchecked += CheckBoxOnUnChecked;
            this.CheckBoxReleaseTipBack2TakePos.Checked += CheckBoxOnChecked;
            this.CheckBoxReleaseTipBack2TakePos.Unchecked += CheckBoxOnUnChecked;
            this.CheckBoxReleaseTipZa0Before.Checked += CheckBoxOnChecked;
            this.CheckBoxReleaseTipZa0Before.Unchecked += CheckBoxOnUnChecked;
            this.CheckBoxReleaseTipZa0After.Checked += CheckBoxOnChecked;
            this.CheckBoxReleaseTipZa0After.Unchecked += CheckBoxOnUnChecked;
            this.TextBoxReleaseTipSpeedCmd.TextChanged += TextBoxOnTextChanged;
            this.TextBoxReleaseTipAfterCmd.TextChanged += TextBoxOnTextChanged;
        }

        private void BtnOnClick(object sender, RoutedEventArgs e)
        {
            if (sender.Equals(this.BtnReleaseTipVariableDistanceStep))
                CmdHelper.Wa(this.mHeadIndex,
                    ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipVariableDistanceStep);
        }

        private void RBtnOnChecked(object sender, RoutedEventArgs e)
        {
            if (sender.Equals(this.RBtnReleaseTipMethodPush))
            {
                this.StackPanelReleaseTipPush.Visibility = Visibility.Visible;
                DataHelper.SaveBool(this.mHeadIndex, true, ref ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipUsePush, () => FileUtils.SaveHead(this.mHeadIndex, ParamsHelper.HeadList[this.mHeadIndex]));
            }
            else if (sender.Equals(this.RBtnReleaseTipMethodHook))
            {
                this.StackPanelReleaseTipPush.Visibility = Visibility.Collapsed;
                DataHelper.SaveBool(this.mHeadIndex, false, ref ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipUsePush, () => FileUtils.SaveHead(this.mHeadIndex, ParamsHelper.HeadList[this.mHeadIndex]));
            }
            else if (sender.Equals(this.RBtnReleaseTipAxisP))
            {
                ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipAxis = EAxis.P;
                FileUtils.SaveHead(this.mHeadIndex, ParamsHelper.HeadList[this.mHeadIndex]);
                this.ControlReleaseTipOffset.SetParam(mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipOffset, ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipAxis);
            }
            else if (sender.Equals(this.RBtnReleaseTipAxisQ))
            {
                ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipAxis = EAxis.Q;
                FileUtils.SaveHead(this.mHeadIndex, ParamsHelper.HeadList[this.mHeadIndex]);
                this.ControlReleaseTipOffset.SetParam(mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipOffset, ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipAxis);
            }
        }

        private void TextBoxOnTextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender.Equals(this.ControlReleaseTipOffset.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, this.ControlReleaseTipOffset.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipOffset);
            else if (sender.Equals(this.TextBoxReleaseTipVariableDistanceStep))
                DataHelper.SaveInt(this.mHeadIndex, this.TextBoxReleaseTipVariableDistanceStep.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipVariableDistanceStep);

            else if (sender.Equals(ControlSettingReleaseTipPosList[0].ControlPrepareReleaseTipPosX.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[0].ControlPrepareReleaseTipPosX.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[0].X);
            else if (sender.Equals(ControlSettingReleaseTipPosList[0].ControlPrepareReleaseTipPosY.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[0].ControlPrepareReleaseTipPosY.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[0].Y);
            else if (sender.Equals(ControlSettingReleaseTipPosList[0].ControlPrepareReleaseTipPosZ.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[0].ControlPrepareReleaseTipPosZ.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[0].Z);
            else if (sender.Equals(ControlSettingReleaseTipPosList[0].ControlReleaseTipPosX.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[0].ControlReleaseTipPosX.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[0].X);
            else if (sender.Equals(ControlSettingReleaseTipPosList[0].ControlReleaseTipPosY.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[0].ControlReleaseTipPosY.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[0].Y);
            else if (sender.Equals(ControlSettingReleaseTipPosList[0].ControlReleaseTipPosZ.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[0].ControlReleaseTipPosZ.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[0].Z);

            else if (sender.Equals(ControlSettingReleaseTipPosList[1].ControlPrepareReleaseTipPosX.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[1].ControlPrepareReleaseTipPosX.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[1].X);
            else if (sender.Equals(ControlSettingReleaseTipPosList[1].ControlPrepareReleaseTipPosY.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[1].ControlPrepareReleaseTipPosY.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[1].Y);
            else if (sender.Equals(ControlSettingReleaseTipPosList[1].ControlPrepareReleaseTipPosZ.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[1].ControlPrepareReleaseTipPosZ.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[1].Z);
            else if (sender.Equals(ControlSettingReleaseTipPosList[1].ControlReleaseTipPosX.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[1].ControlReleaseTipPosX.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[1].X);
            else if (sender.Equals(ControlSettingReleaseTipPosList[1].ControlReleaseTipPosY.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[1].ControlReleaseTipPosY.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[1].Y);
            else if (sender.Equals(ControlSettingReleaseTipPosList[1].ControlReleaseTipPosZ.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[1].ControlReleaseTipPosZ.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[1].Z);

            else if (sender.Equals(ControlSettingReleaseTipPosList[2].ControlPrepareReleaseTipPosX.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[2].ControlPrepareReleaseTipPosX.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[2].X);
            else if (sender.Equals(ControlSettingReleaseTipPosList[2].ControlPrepareReleaseTipPosY.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[2].ControlPrepareReleaseTipPosY.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[2].Y);
            else if (sender.Equals(ControlSettingReleaseTipPosList[2].ControlPrepareReleaseTipPosZ.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[2].ControlPrepareReleaseTipPosZ.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[2].Z);
            else if (sender.Equals(ControlSettingReleaseTipPosList[2].ControlReleaseTipPosX.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[2].ControlReleaseTipPosX.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[2].X);
            else if (sender.Equals(ControlSettingReleaseTipPosList[2].ControlReleaseTipPosY.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[2].ControlReleaseTipPosY.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[2].Y);
            else if (sender.Equals(ControlSettingReleaseTipPosList[2].ControlReleaseTipPosZ.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[2].ControlReleaseTipPosZ.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[2].Z);

            else if (sender.Equals(ControlSettingReleaseTipPosList[3].ControlPrepareReleaseTipPosX.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[3].ControlPrepareReleaseTipPosX.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[3].X);
            else if (sender.Equals(ControlSettingReleaseTipPosList[3].ControlPrepareReleaseTipPosY.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[3].ControlPrepareReleaseTipPosY.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[3].Y);
            else if (sender.Equals(ControlSettingReleaseTipPosList[3].ControlPrepareReleaseTipPosZ.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[3].ControlPrepareReleaseTipPosZ.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipPosList[3].Z);
            else if (sender.Equals(ControlSettingReleaseTipPosList[3].ControlReleaseTipPosX.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[3].ControlReleaseTipPosX.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[3].X);
            else if (sender.Equals(ControlSettingReleaseTipPosList[3].ControlReleaseTipPosY.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[3].ControlReleaseTipPosY.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[3].Y);
            else if (sender.Equals(ControlSettingReleaseTipPosList[3].ControlReleaseTipPosZ.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, ControlSettingReleaseTipPosList[3].ControlReleaseTipPosZ.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosList[3].Z);

            else if (sender.Equals(this.TextBoxReleaseTipSpeedCmd))
                DataHelper.SaveString(this.mHeadIndex, this.TextBoxReleaseTipSpeedCmd.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipSpeedCmd);
            else if (sender.Equals(this.TextBoxReleaseTipAfterCmd))
                DataHelper.SaveString(this.mHeadIndex, this.TextBoxReleaseTipAfterCmd.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipAfterCmd);
        }

        private void ComboBoxOnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender.Equals(this.ComboBoxPushCount))
                DataHelper.SaveInt(this.mHeadIndex, ((int)this.ComboBoxPushCount.SelectedItem).ToString().Trim(), ref ParamsHelper.HeadList[this.mHeadIndex].ReleaseTipUsePushCount, () => FileUtils.SaveHead(this.mHeadIndex, ParamsHelper.HeadList[this.mHeadIndex]));
        }

        private void CheckBoxOnChecked(object sender, RoutedEventArgs e)
        {
            if (sender.Equals(ControlSettingReleaseTipPosList[1].CheckBoxReleaseTipPos))
            {
                ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosAvailableList[1] = true;
                FileUtils.SaveCommonSettings(this.mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex]);
                RefreshReleaseTipPosAvailable(1);
            }
            else if (sender.Equals(ControlSettingReleaseTipPosList[2].CheckBoxReleaseTipPos))
            {
                ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosAvailableList[2] = true;
                FileUtils.SaveCommonSettings(this.mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex]);
                RefreshReleaseTipPosAvailable(2);
            }
            if (sender.Equals(ControlSettingReleaseTipPosList[3].CheckBoxReleaseTipPos))
            {
                ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosAvailableList[3] = true;
                FileUtils.SaveCommonSettings(this.mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex]);
                RefreshReleaseTipPosAvailable(3);
            }

            else if (sender.Equals(this.CheckBoxAxisXGoFirst))
            {
                DataHelper.SaveBool(this.mHeadIndex, true, ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipAxisXGoFirst);
                this.CheckBoxAxisYGoFirst.IsChecked = false;
            }
            else if (sender.Equals(this.CheckBoxAxisYGoFirst))
            {
                DataHelper.SaveBool(this.mHeadIndex, true, ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipAxisYGoFirst);
                this.CheckBoxAxisXGoFirst.IsChecked = false;
            }
            else if (sender.Equals(this.CheckBoxReleaseTipBack2TakePos))
                DataHelper.SaveBool(this.mHeadIndex, true, ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipBack2TakePos);
            else if (sender.Equals(this.CheckBoxReleaseTipZa0Before))
                DataHelper.SaveBool(this.mHeadIndex, true, ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipZa0Before);
            else if (sender.Equals(this.CheckBoxReleaseTipZa0After))
                DataHelper.SaveBool(this.mHeadIndex, true, ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipZa0After);
        }

        private void CheckBoxOnUnChecked(object sender, RoutedEventArgs e)
        {
            if (sender.Equals(ControlSettingReleaseTipPosList[1].CheckBoxReleaseTipPos))
            {
                ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosAvailableList[1] = false;
                FileUtils.SaveCommonSettings(this.mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex]);
                RefreshReleaseTipPosAvailable(1);
            }
            else if (sender.Equals(ControlSettingReleaseTipPosList[2].CheckBoxReleaseTipPos))
            {
                ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosAvailableList[2] = false;
                FileUtils.SaveCommonSettings(this.mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex]);
                RefreshReleaseTipPosAvailable(2);
            }
            if (sender.Equals(ControlSettingReleaseTipPosList[3].CheckBoxReleaseTipPos))
            {
                ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipPosAvailableList[3] = false;
                FileUtils.SaveCommonSettings(this.mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex]);
                RefreshReleaseTipPosAvailable(3);
            }

            else if (sender.Equals(this.CheckBoxAxisXGoFirst))
                DataHelper.SaveBool(this.mHeadIndex, false, ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipAxisXGoFirst);
            else if (sender.Equals(this.CheckBoxAxisYGoFirst))
                DataHelper.SaveBool(this.mHeadIndex, false, ref ParamsHelper.CommonSettingList[this.mHeadIndex].PrepareReleaseTipAxisYGoFirst);
            else if (sender.Equals(this.CheckBoxReleaseTipBack2TakePos))
                DataHelper.SaveBool(this.mHeadIndex, false, ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipBack2TakePos);
            else if (sender.Equals(this.CheckBoxReleaseTipZa0Before))
                DataHelper.SaveBool(this.mHeadIndex, false, ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipZa0Before);
            else if (sender.Equals(this.CheckBoxReleaseTipZa0After))
                DataHelper.SaveBool(this.mHeadIndex, false, ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReleaseTipZa0After);
        }
    }
}
