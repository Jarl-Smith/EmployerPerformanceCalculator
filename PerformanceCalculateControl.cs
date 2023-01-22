using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace EmployerPerformanceCalculator
{
    internal class PerformanceCalculateControl
    {
        public string TargetSubDistrictName {
            get
            {
                return targetSubDistrictName;
            }
            set
            {
                targetSubDistrictName = value;
            }
        }
        public Dictionary<string, string> EmployerPerformanceToMonthPerformanceDict = new Dictionary<string, string>();
        public string OutputExcelPath = @"E:\OfficeWorkBench\123.xlsx";


        private PerformanceExcelHelper monthPerformanceExcelHelper = null;
        private String targetSubDistrictName = "汤阴";


        public void LoadExcel(string monthPerformanceExcelPath)
        {
            if (!File.Exists(monthPerformanceExcelPath))
            {
                return;
            }

            monthPerformanceExcelHelper = new PerformanceExcelHelper();

            monthPerformanceExcelHelper.OpenExcel(monthPerformanceExcelPath, false);
        }

        /// <summary>
        /// 在考核月成绩表中查找考核指标
        /// </summary>
        /// <param name="indicator">需要查找的考核指标</param>
        /// <returns>TotalPoint：该考核的总分，Point：目标县区的得分，Rank：目标县区的名次</returns>
        private dynamic searchPerformanceIndicator(string indicator)
        {
            int totalPointRowNumber = monthPerformanceExcelHelper.FindCellByKeyword("分值").Row;//找到“分值”关键字所在行数

            var cell = monthPerformanceExcelHelper.FindCellByKeyword(indicator);//找到指标所在单元格
            if (cell == null) { return null; }//如果找不到返回null

            double totalPoint = monthPerformanceExcelHelper.FindByRowAndColumnNumber(totalPointRowNumber, cell.Column);//找到该指标的总分值
            var valueFlag = monthPerformanceExcelHelper.FindCellByKeyword(cell.Row, cell.Column, "得分");//找到"得分"关键字所在单元格
            Dictionary<string, double> subdistrictAndValue = new Dictionary<string, double>();
            //依次将所有大区的名称及得分添加到字典中
            for (int i = 1; i < 7; i++)
            {
                string subdistrictName = monthPerformanceExcelHelper.FindByRowAndColumnNumber(valueFlag.Row + i, 1);
                double value1 = monthPerformanceExcelHelper.FindByRowAndColumnNumber(valueFlag.Row + i, valueFlag.Column);
                subdistrictAndValue.Add(subdistrictName, value1);
            }
            //计算县区排名
            int rank = calculateRank(subdistrictAndValue, targetSubDistrictName);
            //将此考核项的总分、得分及排名进行输出
            //Console.WriteLine("{0}\t{1}\t{2}\t{3}", indicator, totalPoint, subdistrictAndValue[targetSubDistrictName], rank);
            return new { TotalPoint = totalPoint, Point = subdistrictAndValue[targetSubDistrictName], Rank = rank };
        }


        public void OutputToEmpExcel()
        {
            PerformanceExcelHelper performanceExcelHelperTemp = new PerformanceExcelHelper();
            performanceExcelHelperTemp.CreateExcel(OutputExcelPath);

            //写入首行标题
            performanceExcelHelperTemp.EditByRowAndColumnNumber(1, 1, "考核项");
            performanceExcelHelperTemp.EditByRowAndColumnNumber(1, 2, "月KPI");
            performanceExcelHelperTemp.EditByRowAndColumnNumber(1, 3, "月得分");
            performanceExcelHelperTemp.EditByRowAndColumnNumber(1, 4, "月排名");

            //开始依次查找并写入考核项、KPI、得分、排名
            int row = 2;
            foreach (string p in EmployerPerformanceToMonthPerformanceDict.Keys)
            {
                performanceExcelHelperTemp.EditByRowAndColumnNumber(row, 1, p);
                var result = searchPerformanceIndicator(EmployerPerformanceToMonthPerformanceDict[p]);
                if (result is not null)
                {
                    performanceExcelHelperTemp.EditByRowAndColumnNumber(row, 2, result.TotalPoint.ToString());
                    performanceExcelHelperTemp.EditByRowAndColumnNumber(row, 3, result.Point.ToString());
                    performanceExcelHelperTemp.EditByRowAndColumnNumber(row, 4, result.Rank.ToString());
                }
                else
                {
                    performanceExcelHelperTemp.EditByRowAndColumnNumber(row, 2, "");
                    performanceExcelHelperTemp.EditByRowAndColumnNumber(row, 3, "");
                    performanceExcelHelperTemp.EditByRowAndColumnNumber(row, 4, "");
                }
                row++;
            }
            performanceExcelHelperTemp.SaveExcel();
            performanceExcelHelperTemp.CloseExel();
        }

        public void Close()
        {
            if (monthPerformanceExcelHelper != null) { monthPerformanceExcelHelper.CloseExel(); }
        }


        /// <summary>
        /// 计算县区排名
        /// </summary>
        /// <param name="subdistrictAndValueDict">键值对形式的县区及得分</param>
        /// <param name="subdistrictName">将要查找排名的县区名称</param>
        /// <returns>县区的排名，返回-1证明找不到</returns>
        private int calculateRank(Dictionary<string, double> subdistrictAndValueDict, string subdistrictName)
        {
            if (!subdistrictAndValueDict.ContainsKey(subdistrictName)) { return -1; }//所给县区名称不在键值对里，证明找不到
            int arraySize = subdistrictAndValueDict.Values.Count;
            double[] values = new double[arraySize];
            subdistrictAndValueDict.Values.CopyTo(values, 0);
            Array.Sort(values);//此处对得分进行排序，排序后是升序
            Array.Reverse(values);//对升序的排名进行反转，变成降序
            int result = Array.IndexOf(values, subdistrictAndValueDict[subdistrictName]);//查找地区的得分所在排名
            return result + 1;//result是索引，索引+1才是名次
        }
    }
}
