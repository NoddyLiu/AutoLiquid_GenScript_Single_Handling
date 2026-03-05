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

namespace AutoLiquid_GenScript_Single_Handling.UserControls
{
    /// <summary>
    /// 只显示盘符
    /// </summary>
    public partial class ControlTemplateOnlyNumber : UserControl
    {
        private int mTemplateIndex;

        public ControlTemplateOnlyNumber(int templateIndex)
        {
            InitializeComponent();

            this.mTemplateIndex = templateIndex;

            InitWidget();
        }

        private void InitWidget()
        {
            this.TextBlockTemplateName.Text = (string) this.FindResource("Template") + (this.mTemplateIndex + 1);
        }
    }
}
