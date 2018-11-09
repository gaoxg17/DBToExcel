using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using DBToExcel.Lib;

namespace DBToExcel
{
    class Program
    {
        private static readonly ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static void Main(string[] args)
        {
            var startMsg = $"开始执行 {DateTime.Now:F}\n";
            Console.WriteLine(startMsg);
            logger.Info("开始执行");

            var eo = new csDBToExcel().fromXmlFile();

            //check config
            var checkConfig = Exec.CheckConfig(eo);
            if (!checkConfig.success)
            {
                Console.WriteLine(checkConfig.errMsg);
                logger.Info(checkConfig.errMsg);
                goto end;
            }
            //to execl
            var toExcel = Exec.ToExcel(eo);
            if (!toExcel.success)
            {
                Console.WriteLine(toExcel.errMsg);
                logger.Info(toExcel.errMsg);
                goto end;
            }

            Console.WriteLine($"本次执行成功\n");
            logger.Info($"本次执行成功\n");

            end: Console.WriteLine("Please enter any key to colse..");
            Console.ReadKey(true);

        }

        static void InitXmlConfig()
        {
            csDBToExcel e2o = new csDBToExcel();
            e2o.Sheets = new List<Sheet>();
            e2o.ConnStr = "connection string";
            e2o.DBType = "oracle";

            Sheet sm = new Sheet();
            sm.SheetName = "sheetname";
            sm.Sql = "sql sentence";

            sm.Fileds = new List<Filed>();

            Filed fm = new Filed();
            fm.FldName = "fieldname";
            fm.RowName = "rowname";

            sm.Fileds.Add(fm);
            e2o.Sheets.Add(sm);

            e2o.toXmlFile();

        }
    }
}
