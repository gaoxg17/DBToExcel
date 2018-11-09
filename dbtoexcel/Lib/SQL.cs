using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Data.SqlClient;

namespace DBToExcel.Lib
{
    public class OracleHelper
    {
        public static string connStr = null;

        #region 执行SQL语句,返回受影响行数
        public static int ExecuteNonQuery(string sql, params OracleParameter[] parameters)
        {
            using (OracleConnection conn = new OracleConnection(connStr))
            {
                conn.Open();
                using (OracleCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.Parameters.AddRange(parameters);
                    return cmd.ExecuteNonQuery();
                }
            }
        }
        #endregion
        #region 执行SQL语句,返回DataTable;只用来执行查询结果比较少的情况
        public static DataTable ExecuteDataTable(string sql, params OracleParameter[] parameters)
        {
            using (OracleConnection conn = new OracleConnection(connStr))
            {
                conn.Open();
                using (OracleCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.Parameters.AddRange(parameters);
                    OracleDataAdapter adapter = new OracleDataAdapter(cmd);
                    DataTable datatable = new DataTable();
                    adapter.Fill(datatable);
                    return datatable;
                }
            }
        }
        #endregion
    }

    public class SqlServerHelper
    {
        public static string connStr = null;

        /// <summary>
        /// 执行增删改操作
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static int ExecuteNonQuery(string sql, params SqlParameter[] para)
        {

            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                //将参数添加到参数集合中
                cmd.Parameters.AddRange(para);
                return cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// 执行查询
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static DataTable ExecuteDataTable(string sql, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();

                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddRange(parameters);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable datatable = new DataTable();
                adapter.Fill(datatable);
                return datatable;
            }
        }
    }

    public class DB
    {
        public static string ORACLE = "oracle";
        public static string SQLSERVER = "sqlserver";
    }
}
