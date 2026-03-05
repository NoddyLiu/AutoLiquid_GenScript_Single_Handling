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
using DataHelper = AutoLiquid_GenScript_Single_Handling.Utils.DataHelper;

namespace AutoLiquid_GenScript_Single_Handling.UserControls
{
    /// <summary>
    /// 体积设置控件
    /// </summary>
    public partial class ControlSettingVolume : UserControl
    {
        // 移液头Index
        private int mHeadIndex;

        public ControlSettingVolume(int headIndex)
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
            this.TextBoxAbsorbAirBeforePercent.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].AbsorbAirBeforePercent.ToString();
            this.TextBoxAirDelayAfterJet.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].AirDelayAfterJet.ToString();
            this.TextBoxAbsorbAirAfterPercent.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].AbsorbAirAfterPercent.ToString();
            // this.TextBoxAbsorbLiquidMorePercent.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].AbsorbLiquidMoreOne2MorePercent.ToString();
            // this.TextBoxJetLiquidMoreScale.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].JetLiquidMoreOne2MoreScale.ToString();
            this.TextBoxAbsorbLiquidMorePercent.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].AbsorbLiquidMorePercent.ToString();
            this.TextBoxReverseJetAfterAbsorbPercent.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].ReverseJetAfterAbsorbPercent.ToString();
            // 多点校准
            this.CheckBoxMultiCalibration.IsChecked = ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.Available;
            if (!ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.Available)
            {
                this.StackPanelMultiCalibration.Visibility = Visibility.Collapsed;
                this.StackPanelMultiCalibrationTest.Visibility = Visibility.Collapsed;
            }
            this.TextBoxPVol1.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PVolArray[0].ToString();
            this.TextBoxPCompensation1.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PCompensationArray[0].ToString();
            this.TextBoxPVol2.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PVolArray[1].ToString();
            this.TextBoxPCompensation2.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PCompensationArray[1].ToString();
            this.TextBoxPVol3.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PVolArray[2].ToString();
            this.TextBoxPCompensation3.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PCompensationArray[2].ToString();
            this.TextBoxPVol4.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PVolArray[3].ToString();
            this.TextBoxPCompensation4.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PCompensationArray[3].ToString();
            this.TextBoxPVol5.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PVolArray[4].ToString();
            this.TextBoxPCompensation5.Text = ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PCompensationArray[4].ToString();
        }

        private void ControlEvent()
        {
            this.TextBoxAbsorbAirBeforePercent.TextChanged += TextBoxOnTextChanged;
            this.TextBoxAirDelayAfterJet.TextChanged += TextBoxOnTextChanged;
            this.TextBoxAbsorbAirAfterPercent.TextChanged += TextBoxOnTextChanged;
            // this.TextBoxAbsorbLiquidMorePercent.TextChanged += TextBoxOnTextChanged;
            // this.TextBoxJetLiquidMoreScale.TextChanged += TextBoxOnTextChanged;
            this.TextBoxAbsorbLiquidMorePercent.TextChanged += TextBoxOnTextChanged;
            this.TextBoxReverseJetAfterAbsorbPercent.TextChanged += TextBoxOnTextChanged;
            // 多点校准
            this.TextBoxPVol1.TextChanged += TextBoxOnTextChanged;
            this.TextBoxPCompensation1.TextChanged += TextBoxOnTextChanged;
            this.TextBoxPVol2.TextChanged += TextBoxOnTextChanged;
            this.TextBoxPCompensation2.TextChanged += TextBoxOnTextChanged;
            this.TextBoxPVol3.TextChanged += TextBoxOnTextChanged;
            this.TextBoxPCompensation3.TextChanged += TextBoxOnTextChanged;
            this.TextBoxPVol4.TextChanged += TextBoxOnTextChanged;
            this.TextBoxPCompensation4.TextChanged += TextBoxOnTextChanged;
            this.TextBoxPVol5.TextChanged += TextBoxOnTextChanged;
            this.TextBoxPCompensation5.TextChanged += TextBoxOnTextChanged;
            this.CheckBoxMultiCalibration.Checked += (sender, args) =>
            {
                ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.Available = true;
                FileUtils.SaveCommonSettings(this.mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex]);
                this.StackPanelMultiCalibration.Visibility = Visibility.Visible;
                this.StackPanelMultiCalibrationTest.Visibility = Visibility.Visible;
            };
            this.CheckBoxMultiCalibration.Unchecked += (sender, args) =>
            {
                ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.Available = false;
                FileUtils.SaveCommonSettings(this.mHeadIndex, ParamsHelper.CommonSettingList[this.mHeadIndex]);
                this.StackPanelMultiCalibration.Visibility = Visibility.Collapsed;
                this.StackPanelMultiCalibrationTest.Visibility = Visibility.Collapsed;
            };
            // 测试
            this.BtnHeadAirSeal.Click += BtnOnClick;
            this.BtnAspVol.Click += BtnOnClick;
            this.BtnDisVol1.Click += BtnOnClick;
            this.BtnDisVol2.Click += BtnOnClick;
            this.BtnDisAll.Click += BtnOnClick;
        }

        private void TextBoxOnTextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender.Equals(this.TextBoxAbsorbAirBeforePercent))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxAbsorbAirBeforePercent.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].AbsorbAirBeforePercent);
            else if (sender.Equals(this.TextBoxAirDelayAfterJet))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxAirDelayAfterJet.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].AirDelayAfterJet);
            else if (sender.Equals(this.TextBoxAbsorbAirAfterPercent))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxAbsorbAirAfterPercent.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].AbsorbAirAfterPercent);
            // else if (sender.Equals(this.TextBoxAbsorbLiquidMorePercent))
            //     DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxAbsorbLiquidMorePercent.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].AbsorbLiquidMoreOne2MorePercent);
            // else if (sender.Equals(this.TextBoxJetLiquidMoreScale))
            //     DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxJetLiquidMoreScale.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].JetLiquidMoreOne2MoreScale);
            else if (sender.Equals(this.TextBoxAbsorbLiquidMorePercent))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxAbsorbLiquidMorePercent.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].AbsorbLiquidMorePercent);
            else if (sender.Equals(this.TextBoxReverseJetAfterAbsorbPercent))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxReverseJetAfterAbsorbPercent.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].ReverseJetAfterAbsorbPercent);
            // 多点校准
            else if (sender.Equals(this.TextBoxPVol1))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxPVol1.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PVolArray[0]);
            else if (sender.Equals(this.TextBoxPCompensation1))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxPCompensation1.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PCompensationArray[0]);
            else if (sender.Equals(this.TextBoxPVol2))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxPVol2.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PVolArray[1]);
            else if (sender.Equals(this.TextBoxPCompensation2))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxPCompensation2.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PCompensationArray[1]);
            else if (sender.Equals(this.TextBoxPVol3))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxPVol3.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PVolArray[2]);
            else if (sender.Equals(this.TextBoxPCompensation3))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxPCompensation3.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PCompensationArray[2]);
            else if (sender.Equals(this.TextBoxPVol4))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxPVol4.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PVolArray[3]);
            else if (sender.Equals(this.TextBoxPCompensation4))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxPCompensation4.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PCompensationArray[3]);
            else if (sender.Equals(this.TextBoxPVol5))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxPVol5.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PVolArray[4]);
            else if (sender.Equals(this.TextBoxPCompensation5))
                DataHelper.SaveDecimal(this.mHeadIndex, this.TextBoxPCompensation5.Text.Trim(), ref ParamsHelper.CommonSettingList[this.mHeadIndex].MultiCalibration.PCompensationArray[4]);
        }

        private void BtnOnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var one2More = this.RBtnJetOne2More.IsChecked.Value;

                if (sender.Equals(this.BtnHeadAirSeal))
                {
                    var vol = decimal.Parse(this.TextBoxHeadAirSeal.Text);
                    CmdHelper.Ps(mHeadIndex, vol, false);
                }
                else if (sender.Equals(this.BtnAspVol))
                {
                    var volOrigin = decimal.Parse(this.TextBoxAspVol.Text);
                    var vol = one2More ? volOrigin : DataHelper.CalibrateVol(this.mHeadIndex, volOrigin);
                    CmdHelper.Ps(mHeadIndex, vol, false);
                }
                else if (sender.Equals(this.BtnDisVol1))
                {
                    if (one2More)
                    {
                        var vol = DataHelper.CalibrateVol(0, decimal.Parse(this.TextBoxDisVol1.Text));
                        CmdHelper.Ps(mHeadIndex, vol * -1, false);
                    }
                    else
                        CmdHelper.Pa(mHeadIndex, 0, false);
                }
                else if (sender.Equals(this.BtnDisVol2))
                {
                    if (one2More)
                    {
                        var vol = DataHelper.CalibrateVol(0, decimal.Parse(this.TextBoxDisVol2.Text));
                        CmdHelper.Ps(mHeadIndex, vol * -1, false);
                    }
                    else
                        CmdHelper.Pa(mHeadIndex, 0, false);
                }
                else if (sender.Equals(this.BtnDisAll))
                {
                    CmdHelper.Pa(mHeadIndex, 0, false);
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message);
                MessageBox.Show(ex.Message, "", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
