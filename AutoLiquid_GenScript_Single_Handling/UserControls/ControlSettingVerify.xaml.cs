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
    /// 验证设置控件
    /// </summary>
    public partial class ControlSettingVerify : UserControl
    {
        // 移液头Index
        private int mHeadIndex;

        public ControlSettingVerify(int headIndex)
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
            this.CheckBoxScanVerify.IsChecked = ParamsHelper.IO.ScanAvailable;
            if (ParamsHelper.IO.ScanPhotoMode)
                this.RBtnModePhoto.IsChecked = true;
            else
                this.RBtnModeVideo.IsChecked = true;
            this.TextBoxTakePhotoDelayMid.Text = ParamsHelper.IO.TakePhotoDelayMid.ToString();
        }

        private void ControlEvent()
        {
            this.CheckBoxScanVerify.Checked += CheckBoxOnChecked;
            this.CheckBoxScanVerify.Unchecked += CheckBoxOnUnChecked;
            this.RBtnModePhoto.Checked += RBtnOnChecked;
            this.RBtnModeVideo.Checked += RBtnOnChecked;
            this.TextBoxTakePhotoDelayMid.TextChanged += TextBoxOnTextChanged;
        }

        private void RBtnOnChecked(object sender, RoutedEventArgs e)
        {
            if (sender.Equals(this.RBtnModePhoto))
                ParamsHelper.IO.ScanPhotoMode= true;
            else if (sender.Equals(this.RBtnModeVideo))
                ParamsHelper.IO.ScanPhotoMode = false;
        }

        private void CheckBoxOnChecked(object sender, RoutedEventArgs e)
        {
            if (sender.Equals(this.CheckBoxScanVerify))
                DataHelper.SaveBool(this.mHeadIndex, true, ref ParamsHelper.IO.ScanAvailable);
        }

        private void CheckBoxOnUnChecked(object sender, RoutedEventArgs e)
        {
            if (sender.Equals(this.CheckBoxScanVerify))
                DataHelper.SaveBool(this.mHeadIndex, false, ref ParamsHelper.IO.ScanAvailable);
        }

        private void TextBoxOnTextChanged(object sender, TextChangedEventArgs e)
        {
           if (sender.Equals(this.TextBoxTakePhotoDelayMid))
                DataHelper.SaveInt(this.mHeadIndex, this.TextBoxTakePhotoDelayMid.Text.Trim(), ref ParamsHelper.IO.TakePhotoDelayMid);
        }
    }
}
