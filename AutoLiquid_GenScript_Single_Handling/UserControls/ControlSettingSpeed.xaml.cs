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
    /// 速度设置控件
    /// </summary>
    public partial class ControlSettingSpeed : UserControl
    {
        // 移液头Index
        private int mHeadIndex;

        public ControlSettingSpeed(int headIndex)
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
            this.TextBoxSpeedXCmd.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].DefaultXSpeed;
            this.TextBoxSpeedXPercent.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].XSpeedPercent.ToString();
            this.TextBoxSpeedXPercent.IsEnabled = !ParamsHelper.CommonSettingList[this.mHeadIndex].DefaultXSpeed.Equals("");
            this.TextBoxSpeedYCmd.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].DefaultYSpeed;
            this.TextBoxSpeedYPercent.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].YSpeedPercent.ToString();
            this.TextBoxSpeedYPercent.IsEnabled = !ParamsHelper.CommonSettingList[this.mHeadIndex].DefaultYSpeed.Equals("");

            this.TextBoxSpeedZCmd.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].DefaultZSpeed;
            this.TextBoxSpeedZPercent.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].ZSpeedPercent.ToString();
            this.TextBoxSpeedZPercent.IsEnabled = !ParamsHelper.CommonSettingList[this.mHeadIndex].DefaultZSpeed.Equals("");
            this.TextBoxSpeedTipLift1Cmd.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].SpeedTipLift1;
            this.TextBoxSpeedTipLift1Percent.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].TipLift1SpeedPercent.ToString();
            this.TextBoxSpeedTipLift1Percent.IsEnabled = !ParamsHelper.CommonSettingList[this.mHeadIndex].SpeedTipLift1.Equals("");
            this.ControlHeightTipLift1.SetParam(this.mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex].TipLift1Height, EAxis.Z);
            this.ControlHeightTipLift1.IsEnabled = !ParamsHelper.CommonSettingList[this.mHeadIndex].SpeedTipLift1.Equals("");

            this.TextBoxSpeedPCmd.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].DefaultPSpeed;
            this.TextBoxSpeedPPercent.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].PSpeedPercent.ToString();
            this.TextBoxSpeedPPercent.IsEnabled = !ParamsHelper.CommonSettingList[this.mHeadIndex].DefaultPSpeed.Equals("");
            this.TextBoxSpeedJet1Cmd.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].SpeedJet1;
            this.TextBoxSpeedJet1Percent.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].Jet1SpeedPercent.ToString();
            this.TextBoxSpeedJet1Percent.IsEnabled = !ParamsHelper.CommonSettingList[this.mHeadIndex].SpeedJet1.Equals("");
            this.TextBoxSpeedJet2Cmd.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].SpeedJet2;
            this.TextBoxSpeedJet2Percent.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].Jet2SpeedPercent.ToString();
            this.TextBoxSpeedJet2Percent.IsEnabled = !ParamsHelper.CommonSettingList[this.mHeadIndex].SpeedJet2.Equals("");
            this.TextBoxVolumeJet2.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].VolumeJet2.ToString();
            this.TextBoxDelayBetweenJet.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].DelayBetweenJet.ToString();
        }

        private void ControlEvent()
        {
            this.TextBoxSpeedXCmd.TextChanged += TextBoxOnTextChanged;
            this.TextBoxSpeedXPercent.TextChanged += TextBoxOnTextChanged;
            this.TextBoxSpeedYCmd.TextChanged += TextBoxOnTextChanged;
            this.TextBoxSpeedYPercent.TextChanged += TextBoxOnTextChanged;

            this.TextBoxSpeedZCmd.TextChanged += TextBoxOnTextChanged;
            this.TextBoxSpeedZPercent.TextChanged += TextBoxOnTextChanged;
            this.TextBoxSpeedTipLift1Cmd.TextChanged += TextBoxOnTextChanged;
            this.TextBoxSpeedTipLift1Percent.TextChanged += TextBoxOnTextChanged;
            this.ControlHeightTipLift1.TextBoxPos.TextChanged += TextBoxOnTextChanged;

            this.TextBoxSpeedPCmd.TextChanged += TextBoxOnTextChanged;
            this.TextBoxSpeedPPercent.TextChanged += TextBoxOnTextChanged;
            this.TextBoxSpeedJet1Cmd.TextChanged += TextBoxOnTextChanged;
            this.TextBoxSpeedJet1Percent.TextChanged += TextBoxOnTextChanged;
            this.TextBoxSpeedJet2Cmd.TextChanged += TextBoxOnTextChanged;
            this.TextBoxSpeedJet2Percent.TextChanged += TextBoxOnTextChanged;
            this.TextBoxVolumeJet2.TextChanged += TextBoxOnTextChanged;
            this.TextBoxDelayBetweenJet.TextChanged += TextBoxOnTextChanged;
        }

        private void TextBoxOnTextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender.Equals(this.TextBoxSpeedXCmd))
            {
                DataHelper.SaveString(this.mHeadIndex, this.TextBoxSpeedXCmd.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].DefaultXSpeed);
                this.TextBoxSpeedXPercent.IsEnabled = !this.TextBoxSpeedXCmd.Text.Equals("");
            }
            else if (sender.Equals(this.TextBoxSpeedXPercent))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxSpeedXPercent.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].XSpeedPercent);
            else if (sender.Equals(this.TextBoxSpeedYCmd))
            {
                DataHelper.SaveString(this.mHeadIndex, this.TextBoxSpeedYCmd.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].DefaultYSpeed);
                this.TextBoxSpeedYPercent.IsEnabled = !this.TextBoxSpeedYCmd.Text.Equals("");
            }
            else if (sender.Equals(this.TextBoxSpeedYPercent))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxSpeedYPercent.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].YSpeedPercent);
            
            else if (sender.Equals(this.TextBoxSpeedZCmd))
            {
                DataHelper.SaveString(this.mHeadIndex, this.TextBoxSpeedZCmd.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].DefaultZSpeed);
                this.TextBoxSpeedZPercent.IsEnabled = !this.TextBoxSpeedZCmd.Text.Equals("");
            }
            else if (sender.Equals(this.TextBoxSpeedZPercent))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxSpeedZPercent.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ZSpeedPercent);
            else if (sender.Equals(this.TextBoxSpeedTipLift1Cmd))
            {
                DataHelper.SaveString(this.mHeadIndex, this.TextBoxSpeedTipLift1Cmd.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].SpeedTipLift1);
                this.TextBoxSpeedTipLift1Percent.IsEnabled = !this.TextBoxSpeedTipLift1Cmd.Text.Equals("");
                this.ControlHeightTipLift1.IsEnabled = !this.TextBoxSpeedTipLift1Cmd.Text.Equals("");
            }
            else if (sender.Equals(this.TextBoxSpeedTipLift1Percent))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxSpeedTipLift1Percent.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].TipLift1SpeedPercent);
            else if (sender.Equals(this.ControlHeightTipLift1.TextBoxPos))
                DataHelper.SaveDecimal(this.mHeadIndex, this.ControlHeightTipLift1.TextBoxPos.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].TipLift1Height);

            else if (sender.Equals(this.TextBoxSpeedPCmd))
            {
                DataHelper.SaveString(this.mHeadIndex, this.TextBoxSpeedPCmd.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].DefaultPSpeed);
                this.TextBoxSpeedPPercent.IsEnabled = !this.TextBoxSpeedPCmd.Text.Equals("");
            }
            else if (sender.Equals(this.TextBoxSpeedPPercent))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxSpeedPPercent.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].PSpeedPercent);

            else if (sender.Equals(this.TextBoxSpeedJet1Cmd))
            {
                DataHelper.SaveString(this.mHeadIndex, this.TextBoxSpeedJet1Cmd.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].SpeedJet1);
                this.TextBoxSpeedJet1Percent.IsEnabled = !this.TextBoxSpeedJet1Cmd.Text.Equals("");
            }
            else if (sender.Equals(this.TextBoxSpeedJet1Percent))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxSpeedJet1Percent.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].Jet1SpeedPercent);
            else if (sender.Equals(this.TextBoxSpeedJet2Cmd))
            {
                DataHelper.SaveString(this.mHeadIndex, this.TextBoxSpeedJet2Cmd.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].SpeedJet2);
                this.TextBoxSpeedJet2Percent.IsEnabled = !this.TextBoxSpeedJet2Cmd.Text.Equals("");
            }
            else if (sender.Equals(this.TextBoxSpeedJet2Percent))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxSpeedJet2Percent.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].Jet2SpeedPercent);
            else if (sender.Equals(this.TextBoxVolumeJet2))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxVolumeJet2.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].VolumeJet2);
            else if (sender.Equals(this.TextBoxDelayBetweenJet))
                DataHelper.SaveInt(this.mHeadIndex, this.TextBoxDelayBetweenJet.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].DelayBetweenJet);
        }
    }
}
