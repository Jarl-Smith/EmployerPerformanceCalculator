using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmployerPerformanceCalculator
{
    internal class PerformanceExcelReader
    {
        private String employerExcel;
        private String monthPerformanceExcel;


        public string EmployerExcel { get => employerExcel; set => employerExcel = value; }
        public string PerformanceExcel { get => monthPerformanceExcel; set => monthPerformanceExcel = value; }

        public void SetExcels(string employerExcelPath,string monthPerformanceExcelPath)
        {
            if(!Directory.Exists(employerExcelPath) || !Directory.Exists(monthPerformanceExcelPath))//如果两个文件只要有一个不存在，结束方法
            {
                return;
            }
            //TODO执行读取文件的方法，并对属性进行赋值

        }

        private void loadExcelFile(string fileName)
        {

        }

        /// <summary>
        /// 计算县区排名
        /// </summary>
        /// <param name="subDistrictAndValueDict">键值对形式的县区及得分</param>
        /// <param name="subDistrictName">将要查找排名的县区名称</param>
        /// <returns>县区的排名，返回-1证明找不到</returns>
        internal int caculateRank(Dictionary<string, double> subDistrictAndValueDict, string subDistrictName)
        {
            if (subDistrictAndValueDict.ContainsKey(subDistrictName)) { return -1; }//所给县区名称不在键值对里
            int arraySize = subDistrictAndValueDict.Values.Count;
            double[] values = new double[arraySize];
            subDistrictAndValueDict.Values.CopyTo(values, 0);
            Array.Sort(values);//此处对得分进行排序，排序后是升序
            Array.Reverse(values);//对升序的排名进行反转，变成降序
            int result = Array.IndexOf(values, subDistrictAndValueDict[subDistrictName]);//查找地区的得分所在排名
            if (result == -1) { return 0; }//如果上边返回-1，证明找不到
            return result + 1;//如果找到了，返回名次，result是索引，索引+1才是名次
        }
    }
}
