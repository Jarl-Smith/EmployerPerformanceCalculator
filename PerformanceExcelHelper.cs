using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace EmployerPerformanceCalculator
{
    internal class PerformanceExcelHelper : IDisposable
    {
        private Excel.Application oXL = null;//声明Excel Application对象
        private Excel.Workbook oWB = null;//声明工作薄对象
        private Excel.Worksheet oSheet = null;//声明工作表对象
        private Excel.Range oRng = null;//声明选中单元格的对象

        public void CreateExcel(string filePath)
        {
            oXL = new Excel.Application();
            oWB = oXL.Workbooks.Add();
            oSheet = oWB.ActiveSheet;
            oWB.SaveAs(filePath);
        }

        /// <summary>
        /// 打开excel文件，设置字段成员变量
        /// </summary>
        /// <param name="path">excel文件路径</param>
        public void OpenExcel(string path, bool visable = true)
        {
            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = visable;

                oWB = oXL.Workbooks.Open(path);//传递文件路径，打开excel文件
                oSheet = (Excel.Worksheet)oWB.ActiveSheet;//获取正在活动的工作表

            }
            catch (Exception theException)
            {
                CloseExel();
                String errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
            }
        }

        /// <summary>
        /// 查找关键字所在单元格,本方法只返回第一次查找到的结果
        /// </summary>
        /// <param name="keyword">关键字</param>
        /// <returns>返回含有Row、Column字段的匿名类，Row为单元格行号，Column为单元格列号，null为未找到关键字</returns>
        public dynamic FindCellByKeyword(string keyword)
        {
            #region 以下代码摘自官网
            ////当搜索到达指定的搜索区域末尾时，它会绕到该区域开头位置。 若要在发生此绕回时停止搜索，请保存第一个找到的单元格的地址，然后针对此保存的地址测试每个连续找到的单元格地址。
            //Excel.Range currentFind = null;
            //Excel.Range firstFind = null;
            //oRng = oSheet.Cells.EntireRow;//将查找范围定义为整个表格
            //currentFind = oRng.Find(keyword,
            //    Missing.Value, Excel.XlFindLookIn.xlValues,
            //    Excel.XlLookAt.xlPart,
            //    Excel.XlSearchOrder.xlByRows,
            //    Excel.XlSearchDirection.xlNext,
            //    false,
            //    Missing.Value,
            //    Missing.Value);
            //while (currentFind != null)
            //{
            //    if (firstFind == null)
            //    {
            //        firstFind = currentFind;
            //    }
            //    else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1) == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
            //    {
            //        break;
            //    }
            //    //这里开始执行找到单元格后需要做的事情
            //    Console.WriteLine("{0}在第{1}行,第{2}列", "1月", currentFind.Row, currentFind.Column);

            //    //按照原来的设置参数继续查找
            //    currentFind = oRng.FindNext(currentFind);
            //}
            #endregion

            oRng = oSheet.Cells.EntireRow;//将查找范围定义为整个表格
            Excel.Range currentFind = oRng.Find(keyword,
                                                Missing.Value,
                                                Excel.XlFindLookIn.xlValues,
                                                Excel.XlLookAt.xlPart,
                                                Excel.XlSearchOrder.xlByRows,
                                                Excel.XlSearchDirection.xlNext,
                                                false,
                                                Missing.Value,
                                                Missing.Value);
            if (currentFind != null)
            {
                return new { Row = currentFind.Row, Column = currentFind.Column };
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 通过指定某个单元格，以此单元格开始查找关键字，返回第一个查找到的结果
        /// </summary>
        /// <param name="searchStartRow">行号</param>
        /// <param name="searchStartColumn">列号</param>
        /// <param name="keyword">关键字</param>
        /// <returns>返回含有Row、Column字段的匿名类，Row为单元格行号，Column为单元格列号，null为未找到关键字</returns>
        public dynamic FindCellByKeyword(int searchStartRow, int searchStartColumn, string keyword)
        {
            oRng = oSheet.Cells.EntireRow;//将查找范围定义为整个表格
            Excel.Range currentFind = oRng.Find(keyword,
                                                oSheet.Cells[searchStartRow, searchStartColumn],
                                                Excel.XlFindLookIn.xlValues,
                                                Excel.XlLookAt.xlPart,
                                                Excel.XlSearchOrder.xlByColumns,
                                                Excel.XlSearchDirection.xlNext,
                                                false,
                                                Missing.Value,
                                                Missing.Value);
            if (currentFind != null)
            {
                return new { Row = currentFind.Row, Column = currentFind.Column };
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 通过行号、列号查找单元格保存的内容
        /// </summary>
        /// <param name="row">行号</param>
        /// <param name="column">列号</param>
        /// <returns>该单元格保存的内容</returns>
        public dynamic FindByRowAndColumnNumber(int row, int column)
        {
            return oSheet.Cells[row, column].Value2;
        }

        /// <summary>
        /// 通过指定行号、列号修改单元格内容
        /// </summary>
        /// <param name="row">行号</param>
        /// <param name="column">列号</param>
        /// <param name="content">需要修改后的内容</param>
        public void EditByRowAndColumnNumber(int row, int column, string content)
        {
            oSheet.Cells[row, column] = content;
        }

        public void SaveExcel()
        {
            oWB.Save();
        }

        /// <summary>
        /// 关闭excel文件，释放进程
        /// </summary>
        public void CloseExel()
        {
            if (oRng != null)
            {
                oRng.Clear();
            }

            if (oWB != null)
            {
                oWB.Close(false);
            }
            if (oXL != null)
            {
                oXL.Quit();
            }

            Dispose();
        }

        /// <summary>
        /// 实现IDisposable接口
        /// </summary>
        public void Dispose()
        {
            if (oRng != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRng);
                oRng = null;
            }
            if (oSheet != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet);
                oSheet = null;
            }
            if (oWB != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWB);
                oWB = null;
            }
            if (oXL != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                oXL = null;
            }
            GC.Collect();
        }
    }
}
