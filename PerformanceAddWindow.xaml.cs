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
using System.Windows.Shapes;

namespace EmployerPerformanceCalculator
{
    /// <summary>
    /// PerformanceAddWindow.xaml 的交互逻辑
    /// </summary>
    public partial class PerformanceAddWindow : Window
    {
        private MyPerformanceCollection mainWindowColl;
        public PerformanceAddWindow(MyPerformanceCollection keyValuePairs)
        {
            InitializeComponent();
            mainWindowColl = keyValuePairs;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            mainWindowColl.Add(new KeyValuePair<string, string>(sourcePerformance.Text, targetPerformance.Text));
            this.Close();
        }
    }
}
