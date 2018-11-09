using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using excel = Microsoft.Office.Interop.Excel;
using System.Data;
using log4net;
using System.Collections;
using System.Data.OleDb;

namespace DBToExcel.Lib
{
    public class Exec
    {
        /// <summary>
        /// 私有日志对象
        /// </summary>
        private static readonly ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static (bool success, string errMsg) CheckConfig(csDBToExcel eo)
        {
            if (eo == null)
                return (false, "配置错误");
            if (eo.DBType != DB.ORACLE && eo.DBType != DB.SQLSERVER)
                return (false, "数据库类型错误");
            if (string.IsNullOrWhiteSpace(eo.ConnStr))
                return (false, "连接字符串错误"); ;
            if (eo.Sheets == null || eo.Sheets.Count == 0)
                return (false, "Sheet节点存在错误");
            foreach (var sheet in eo.Sheets)
            {
                if (string.IsNullOrWhiteSpace(sheet.SheetName))
                    return (false, "SheetName不能为空");
                if (sheet.Fileds == null || sheet.Fileds.Count == 0)
                    return (false, "Filed节点存在错误");
                foreach (var filed in sheet.Fileds)
                {
                    if (string.IsNullOrEmpty(filed.FldName))
                        return (false, "FldName不能为空");
                    if (string.IsNullOrEmpty(filed.RowName))
                        return (false, "配置错误");
                }
            }
            return (true, null);
        }

        public static (bool success, string errMsg) ToExcel(csDBToExcel eo)
        {
            return eo.DBType == DB.ORACLE ? OracleToExcel(eo) : SqlServerToExcel(eo);
        }

        private static (bool success, string errMsg) OracleToExcel(csDBToExcel eo)
        {
            var errMsg = string.Empty;
            try
            {
                var outpath = System.AppDomain.CurrentDomain.BaseDirectory + "outpath";
                if (!Directory.Exists(outpath))
                {
                    Directory.CreateDirectory(outpath);
                }

                OracleHelper.connStr = eo.ConnStr;
                List<Sheet> sheets = new List<Sheet>();
                List<DataTable> dts = new List<DataTable>();
                eo.Sheets.ForEach(e =>
                {
                    try
                    {
                        sheets.Add(e);
                        dts.Add(OracleHelper.ExecuteDataTable(e.Sql));
                    }
                    catch
                    {
                        errMsg += $"Sheet {e.SheetName} 查询异常\n";
                        logger.Error("异常描述：\t" + errMsg);
                    }
                });
                if (string.IsNullOrEmpty(errMsg))
                {
                    ToExcel(sheets, dts, outpath + $"\\{DateTime.Now:D}.xlsx");
                }
            }
            catch (Exception ex)
            {
                logger.Error("异常描述：\t" + ex.Message);
                logger.Error("异常方法：\t" + ex.TargetSite);
                logger.Error("异常堆栈：\t" + ex.StackTrace);
                errMsg += "生成Excel异常\n";
            }
            return (string.IsNullOrEmpty(errMsg), errMsg);
        }

        private static (bool success, string errMsg) SqlServerToExcel(csDBToExcel eo)
        {
            var errMsg = string.Empty;
            try
            {
                var outpath = System.AppDomain.CurrentDomain.BaseDirectory + "outpath";
                if (!Directory.Exists(outpath))
                {
                    Directory.CreateDirectory(outpath);
                }

                SqlServerHelper.connStr = eo.ConnStr;
                List<Sheet> sheets = new List<Sheet>();
                List<DataTable> dts = new List<DataTable>();
                eo.Sheets.ForEach(e =>
                {
                    try
                    {
                        sheets.Add(e);
                        dts.Add(SqlServerHelper.ExecuteDataTable(e.Sql));
                    }
                    catch
                    {
                        errMsg += $"Sheet {e.SheetName} 查询异常\n";
                        logger.Error("异常描述：\t" + errMsg);
                    }
                });
                if (string.IsNullOrEmpty(errMsg))
                {
                    ToExcel(sheets, dts, outpath + $"\\{DateTime.Now:D}.xlsx");
                }
            }
            catch (Exception ex)
            {
                logger.Error("异常描述：\t" + ex.Message);
                logger.Error("异常方法：\t" + ex.TargetSite);
                logger.Error("异常堆栈：\t" + ex.StackTrace);
                errMsg += "生成Excel异常\n";
            }
            return (string.IsNullOrEmpty(errMsg), errMsg);
        }

        private static void ToExcel(List<Sheet> sheets, List<DataTable> dts, string file)
        {
            excel.Application appexcel = null;
            excel.Workbook workbookdata = null;
            excel.Worksheet worksheetdata = null;
            try
            {
                appexcel = new excel.Application();
                workbookdata = appexcel.Workbooks.Add();
                //设置对象不可见  
                appexcel.Visible = false;
                appexcel.DisplayAlerts = false;
                System.Reflection.Missing miss = System.Reflection.Missing.Value;

                int sheetNum = 0;
                foreach (var sheet in sheets)
                {
                    DataTable dt = dts[sheetNum];
                    worksheetdata = (excel.Worksheet)workbookdata.Worksheets.Add(miss, workbookdata.ActiveSheet);
                    //给工作表赋名称  
                    worksheetdata.Name = sheet.SheetName;
                    for (int i = 0; i < sheet.Fileds.Count; i++)
                    {
                        worksheetdata.Cells[1, i + 1] = sheet.Fileds[i].RowName;
                    }

                    //irowcount为实际行数，最大行  
                    int irowcount = dt.Rows.Count;
                    //已执行的行数
                    int iparstedrow = 0;
                    //单次执行的条数
                    int icurrsize = 0;

                    //ieachsize为每次写行的数值，可以自己设置  
                    int ieachsize = 10000;

                    //icolumnaccount为实际列数，最大列数  
                    int icolumnaccount = sheet.Fileds.Count;

                    //在内存中声明一个ieachsize×icolumnaccount的数组，ieachsize是每次最大存储的行数，icolumnaccount就是存储的实际列数  
                    object[,] objval = new object[ieachsize, icolumnaccount];
                    icurrsize = ieachsize;
                    while (iparstedrow < irowcount)
                    {
                        if ((irowcount - iparstedrow) < ieachsize)
                            icurrsize = irowcount - iparstedrow;

                        //用for循环给数组赋值  
                        for (int i = 0; i < icurrsize; i++)
                        {
                            DataRow row = dt.Rows[iparstedrow + i];
                            for (int j = 0; j < icolumnaccount; j++)
                            {
                                if (dt.Columns.Contains(sheet.Fileds[j].FldName))
                                {
                                    var v = row[sheet.Fileds[j].FldName];
                                    objval[i, j] = v != null ? v.ToString() : "";
                                }
                                else
                                {
                                    objval[i, j] = "N/A";
                                }
                            }
                        }
                        string X = "A" + ((int)(iparstedrow + 2)).ToString(); //因为第一行已经写了表头，所以所有数据都应该从a2开始 
                        string col = "";
                        if (icolumnaccount <= 26)
                        {
                            col = ((char)('A' + icolumnaccount - 1)).ToString() + ((int)(iparstedrow + icurrsize + 1)).ToString();
                        }
                        else
                        {
                            col = ((char)('A' + (icolumnaccount / 26 - 1))).ToString() + ((char)('A' + (icolumnaccount % 26 - 1))).ToString() + ((int)(iparstedrow + icurrsize + 1)).ToString();
                        }
                        excel.Range xlrang = worksheetdata.get_Range(X, col);
                        xlrang.NumberFormat = "@";
                        //调用range的value2属性，把内存中的值赋给excel  
                        xlrang.Value2 = objval;
                        iparstedrow = iparstedrow + icurrsize;
                    }
                    sheetNum++;
                }

                ((excel.Worksheet)workbookdata.Worksheets["Sheet1"]).Delete();
                ((excel.Worksheet)workbookdata.Worksheets["Sheet2"]).Delete();
                ((excel.Worksheet)workbookdata.Worksheets["Sheet3"]).Delete();
                //保存工作表  
                workbookdata.SaveAs(file, miss, miss, miss, miss, miss, excel.XlSaveAsAccessMode.xlNoChange, miss, miss, miss);
                workbookdata.Close(false, miss, miss);
                appexcel.Workbooks.Close();
                appexcel.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbookdata);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(appexcel.Workbooks);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(appexcel);
                GC.Collect();
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
            }
        }
    }
}
