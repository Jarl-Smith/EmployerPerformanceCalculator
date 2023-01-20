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

namespace EmployerPerformanceCalculator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private PerformanceCalculateControl calculateControl = new PerformanceCalculateControl();

        private string monthPerformanceFile;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void onChooseMonthPerformanceBtnClick(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Filter by Excel Worksheets
            dlg.Filter = "Excel Worksheets|*.xls;*.xlsx";

            // Show open file dialog box
            bool? result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                monthPerformanceFile = dlg.FileName;
                tbk2.Text = monthPerformanceFile;
            }
        }

        private void onCalculateBtnClick(object sender, RoutedEventArgs e)
        {
            //设置目标县区
            calculateControl.TargetSubDistrictName = subDistrictCmb.Text;
            //从静态资源中找到用户所输入的所有考核项目
            MyPerformanceCollection myColl = (MyPerformanceCollection)this.Resources["myPerformanceColl"];

            //将考核集合转换为字典并赋值
            calculateControl.EmployerPerformanceToMonthPerformanceDict = new Dictionary<string, string>(myColl);
            //计算考核并导出到文件中
            calculateControl.LoadExcel(monthPerformanceFile);
            calculateControl.OutputToEmpExcel();
        }

        private void onAddPerformanceItemBtnClick(object sender, RoutedEventArgs e)
        {
            MyPerformanceCollection temp = (MyPerformanceCollection)this.Resources["myPerformanceColl"];
            PerformanceAddWindow performanceAddWindow = new PerformanceAddWindow(temp);
            performanceAddWindow.Show();
        }

        private void onSubtractPerformanceBtnClick(object sender, RoutedEventArgs e)
        {
            KeyValuePair<string, string> temp = (KeyValuePair<string, string>)performanceCollLsB.SelectedItem;
            MyPerformanceCollection myColl = (MyPerformanceCollection)this.Resources["myPerformanceColl"];
            myColl.Remove(temp);
        }
    }
}
