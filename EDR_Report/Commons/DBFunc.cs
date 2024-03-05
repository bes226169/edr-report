using Dapper;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.Reflection;

namespace EDR_Report
{
    /// <summary>
    /// DBFunc 簡化版本
    /// </summary>
    public class DBFunc
    {
        /// <summary>
        /// 設定檔名稱，預設為appsettings.json。
        /// </summary>
        public static string AppSettingsFile { get; private set; } = "appsettings.json";
        static IConfiguration? conf { get; set; }
        /// <summary>
        /// DBFunc 簡化版本
        /// </summary>
        public DBFunc() 
        {
            conf = new ConfigurationBuilder().AddJsonFile(AppSettingsFile).Build();
        }
        /// <summary>
        /// 初始化，可自訂設定檔名稱。
        /// </summary>
        /// <param name="appSettingsFile"></param>
        public DBFunc(string appSettingsFile)
        {
            AppSettingsFile = appSettingsFile;
            conf = new ConfigurationBuilder().AddJsonFile(AppSettingsFile).Build();
        }
        /// <summary>
        /// 查詢
        /// </summary>
        /// <typeparam name="T">如果沒有指定 Model , 請輸入 dynamic</typeparam>
        /// <param name="connstr"></param>
        /// <param name="sql"></param>
        /// <param name="paramObjs"></param>
        /// <returns></returns>
        public IEnumerable<T> query<T>(string connstr, string sql, object? paramObjs = null)
        {
            OracleConnection? db = null;
            try
            {
                db = new(GetConnectionString(connstr));
                var rt = db.Query<T>(sql, paramObjs);
                db?.Close();
                return rt;
            }
            catch
            {
                db?.Close();
                throw;
            }
        }
        /// <summary>
        /// 執行 SQL Script
        /// </summary>
        /// <param name="connstr"></param>
        /// <param name="sql"></param>
        /// <param name="paramObjs"></param>
        /// <returns></returns>
        public int exec(string connstr, string sql, object? paramObjs = null)
        {
            OracleConnection? db = null;
            try
            {
                db = new(GetConnectionString(connstr));
                var rt = db.Execute(sql, paramObjs);
                db?.Close();
                return rt;
            }
            catch
            {
                db?.Close();
                throw;
            }
        }
        /// <summary>
        /// 執行Stored Procedure
        /// </summary>
        /// <param name="connstr"></param>
        /// <param name="sp"></param>
        /// <param name="paramObjs"></param>
        /// <returns></returns>
        public int sp(string connstr, string sp, DynamicParameters? paramObjs = null)
        {
            OracleConnection? db = null;
            try
            {
                db = new(GetConnectionString(connstr));
                var rt = db.Execute(
                    sql: sp,
                    param: paramObjs,
                    commandType: CommandType.StoredProcedure
                    );
                db?.Close();
                return rt;
            }
            catch
            {
                db?.Close();
                throw;
            }
        }
        /// <summary>
        /// 執行Insert（一筆）。
        /// <para>這個方法可以防止 SQL Injection</para>
        /// </summary>
        /// <param name="connstr">設定檔ConnectionStrings對應的Key</param>
        /// <param name="table">Table Name</param>
        /// <param name="param">欄位及參數</param>
        /// <returns></returns>
        public int Insert(string connstr, string table, object param)
        {
            (var sql, var dp) = InsertSqlWithParams(table, param);
            return exec(connstr, sql.TrimEnd(';'), dp);
        }
        /// <summary>
        /// 產生Insert的SQL Script（一筆）及參數（DynamicParameters）。
        /// <para>這個方法可以防止 SQL Injection</para>
        /// </summary>
        /// <param name="table">Table Name</param>
        /// <param name="param">欄位及對應的值</param>
        /// <returns></returns>
        public (string sql, DynamicParameters dParams) InsertSqlWithParams(string table, object param)
        {
            var li = new List<string>();
            var dp = new DynamicParameters();
            void addParams(string p, object? v)
            {
                var dbt = GetDbType(v);
                if (dbt == null) dp.Add(p, v);
                else dp.Add(p, v, dbt);
            }
            if (param is DynamicParameters)
            {
                foreach (var p in dp.ParameterNames)
                {
                    var v = dp.Get<object?>(p);
                    li.Add(p);
                    addParams(p, v);
                }
            }
            else if (param is IDictionary<string, object?> dic)
            {
                foreach (var kvp in dic)
                {
                    li.Add(kvp.Key);
                    addParams(kvp.Key, kvp.Value);
                }
            }
            else
            {
                var type = param.GetType();
                foreach (PropertyInfo p in type.GetProperties())
                {
                    var v = p.GetValue(param);
                    li.Add(p.Name);
                    addParams(p.Name, v);
                }
            }
            var sql = $"INSERT INTO {table} ({string.Join(',', li)}) VALUES (:{string.Join(", :", li)});";
            return (sql, dp);
        }
        /// <summary>
        /// 組合 SQL Script: Insert
        /// </summary>
        /// <param name="table"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public string insertSql(string table, object param) => InsertSql(table, param);
        /// <summary>
        /// 組合 SQL Script: Insert
        /// </summary>
        /// <param name="table"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public static string InsertSql(string table, object param) => InsertSqls(table, new object[] { param });
        /// <summary>
        /// 組合 SQL Script: Insert (多筆)
        /// </summary>
        /// <param name="table"></param>
        /// <param name="paramObjs"></param>
        /// <returns></returns>
        public static string InsertSqls(string table, object[] paramObjs)
        {
            var rt = string.Empty;
            foreach (var p in paramObjs)
            {
                (var cols, var vals) = ParamsToInsertString(p);
                rt += $"INSERT INTO {table} ({cols}) VALUES ({vals});";
            }
            return rt;

        }
        /// <summary>
        /// 執行Update。
        /// <para>這個方法可以防止 SQL Injection</para>
        /// </summary>
        /// <param name="connstr">設定檔ConnectionStrings對應的Key</param>
        /// <param name="table">Table Name</param>
        /// <param name="setParam">要更新的欄位及值</param>
        /// <param name="whereParam">條件欄位及值</param>
        /// <returns></returns>
        public int Update(string connstr, string table, object setParam, object? whereParam = null)
        {
            (var sql, var dp) = UpdateSqlWithParams(table, setParam, whereParam);
            return exec(connstr, sql.TrimEnd(';'), dp);
        }
        /// <summary>
        /// 產生Update的SQL Script（一筆）及參數（DynamicParameters）。
        /// <para>這個方法可以防止 SQL Injection</para>
        /// </summary>
        /// <param name="table">Table Name</param>
        /// <param name="setParam">要更新的欄位及值</param>
        /// <param name="whereParam">條件欄位及值</param>
        /// <returns></returns>
        public (string sql, DynamicParameters dParams) UpdateSqlWithParams(string table, object setParam, object? whereParam = null)
        {
            var sets = new List<string>();
            var wheres = new List<string>();
            var dp = new DynamicParameters();
            void addParams(string p, object? v)
            {
                var dbt = GetDbType(v);
                if (dbt == null) dp.Add(p, v);
                else dp.Add(p, v, dbt);
            }
            void build(object? param, ref List<string> li)
            {
                if (param == null) return;
                if (param is DynamicParameters)
                {
                    foreach (var p in dp.ParameterNames)
                    {
                        li.Add($"{p} = :{p}");
                        var v = dp.Get<object?>(p);
                        addParams(p, v);
                    }
                }
                else if (param is IDictionary<string, object?> dic)
                {
                    foreach (var kvp in dic)
                    {
                        li.Add($"{kvp.Key} = :{kvp.Key}");
                        addParams(kvp.Key, kvp.Value);
                    }
                }
                else
                {
                    var type = param.GetType();
                    foreach (PropertyInfo p in type.GetProperties())
                    {
                        var v = p.GetValue(param);
                        li.Add($"{p.Name} = :{p.Name}");
                        addParams(p.Name, v);
                    }
                }
            }
            build(setParam, ref sets);

            var sql = $"UPDATE {table} SET {string.Join(',', sets)}";
            if (whereParam != null)
            {
                build(whereParam, ref wheres);
                sql += $" WHERE {string.Join(" AND ", wheres)};";
            }
            return (sql, dp);
        }
        /// <summary>
        /// 組合 SQL Script: Update
        /// </summary>
        /// <param name="table"></param>
        /// <param name="setParam"></param>
        /// <param name="whereParam"></param>
        /// <returns></returns>
        public string updateSql(string table, object setParam, object? whereParam = null) => UpdateSql(table, setParam, whereParam);
        /// <summary>
        /// 組合 SQL Script: Update
        /// </summary>
        /// <param name="table"></param>
        /// <param name="setParam"></param>
        /// <param name="whereParam"></param>
        /// <returns></returns>
        public static string UpdateSql(string table, object setParam, object? whereParam = null)
        {
            var rt = $"UPDATE {table} SET {ParamsToString(setParam)}";
            string where = string.Empty;
            if (whereParam != null) where = ParamsToString(whereParam, " AND ");
            if (!string.IsNullOrEmpty(where)) rt += $" WHERE {where}";
            return rt + ";";
        }
        /// <summary>
        /// 組合 SQL Script: Delete
        /// </summary>
        /// <param name="table"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public string deleteSql(string table, object? param = null) => DeleteSql(table, param);
        /// <summary>
        /// 組合 SQL Script: Delete
        /// </summary>
        /// <param name="table"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public static string DeleteSql(string table, object? param = null)
        {
            var where = string.Empty;
            if (param != null) where = ParamsToString(param, " AND ");
            if (!string.IsNullOrEmpty(where)) where = $" WHERE {where}";
            return $"DELETE {table}{where};";
        }
        public static DataTable ToDataTable(IEnumerable<object> items)
        {
            var data = items.ToArray();
            if (data.Count() == 0) return new DataTable();
            var dt = new DataTable();
            var ColumnsCnt = 1;
            foreach (string key in ((IDictionary<string, object>)data[0]).Keys)
            {
                dt.Columns.Add(dt.Columns.Contains(key) ? $"{key}_{ColumnsCnt++}" : key.ToString());
            }
            foreach (object d in data)
            {
                dt.Rows.Add(((IDictionary<string, object>)d).Values.ToArray());
            }
            return dt;
        }
        /// <summary>
        /// 執行Delete。
        /// <para>這個方法可以防止 SQL Injection</para>
        /// </summary>
        /// <param name="connstr">設定檔ConnectionStrings對應的Key</param>
        /// <param name="table">Table Name</param>
        /// <param name="param">刪除條件</param>
        /// <returns></returns>
        public int Delete(string connstr, string table, object? param = null)
        {
            (var sql, var dp) = DeleteSqlWithParams(table, param);
            return exec(connstr, sql.TrimEnd(';'), dp);
        }
        /// <summary>
        /// 產生Delete的SQL Script（一筆）及參數（DynamicParameters）。
        /// <para>這個方法可以防止 SQL Injection</para>
        /// </summary>
        /// <param name="table">Table Name</param>
        /// <param name="param">刪除條件</param>
        /// <returns></returns>
        public (string sql, DynamicParameters dParams) DeleteSqlWithParams(string table, object? param = null)
        {
            var li = new List<string>();
            var dp = new DynamicParameters();
            void addParams(string p, object? v)
            {
                var dbt = GetDbType(v);
                if (dbt == null) dp.Add(p, v);
                else dp.Add(p, v, dbt);
            }
            var sql = $"DELETE {table}";
            if (param != null)
            {
                if (param is DynamicParameters)
                {
                    foreach (var p in dp.ParameterNames)
                    {
                        li.Add($"{p} = :{p}");
                        var v = dp.Get<object?>(p);
                        addParams(p, v);
                    }
                }
                else if (param is IDictionary<string, object?> dic)
                {
                    foreach (var kvp in dic)
                    {
                        li.Add($"{kvp.Key} = :{kvp.Key}");
                        addParams(kvp.Key, kvp.Value);
                    }
                }
                else
                {
                    var type = param.GetType();
                    foreach (PropertyInfo p in type.GetProperties())
                    {
                        var v = p.GetValue(param);
                        li.Add($"{p.Name} = :{p.Name}");
                        addParams(p.Name, v);
                    }
                }
                sql += $" WHERE {string.Join(" AND ", li)};";
            }
            return (sql, dp);
        }
        /// <summary>
        /// 判斷資料類型
        /// </summary>
        /// <param name="v"></param>
        /// <returns></returns>
        static string SetValueString(object? v)
        {
            if (v == null) return "NULL";
            else if (v is DateTime time) return $"TO_DATE('{time:yyyy-MM-dd HH:mm:ss}', 'yyyy-mm-dd hh24:mi:ss')";
            else if (v is DateTimeOffset dtoff) return $"TO_DATE('{dtoff:yyyy-MM-dd HH:mm:ss}', 'yyyy-mm-dd hh24:mi:ss')";
            else if (v is bool b) return $"'{(b ? "1" : "0")}'";
            else if (v is decimal or byte or short or int or long or double or float) return v.ToString()!;
            else return $"N'{Convert.ToString(v)!.Replace("'", "''")}'";
        }
        /// <summary>
        /// Insert 用, 組合 Columns 及 Values
        /// </summary>
        /// <param name="param"></param>
        /// <returns></returns>
        static (string cols, string vals) ParamsToInsertString(object? param)
        {
            if (param == null) return (string.Empty, string.Empty);
            List<string> cl = new();
            List<string> vl = new();

            if (param is DynamicParameters dp)
            {
                foreach (string p in dp.ParameterNames)
                {
                    cl.Add(p);
                    vl.Add(SetValueString(dp.Get<object?>(p)));
                }
            }
            else if (param is IDictionary<string, object?> dic)
            {
                foreach (var kvp in dic)
                {
                    cl.Add(kvp.Key);
                    vl.Add(SetValueString(kvp.Value));
                }
            }
            else
            {
                var type = param.GetType();
                foreach (PropertyInfo p in type.GetProperties())
                {
                    cl.Add(p.Name);
                    vl.Add(SetValueString(p.GetValue(param)));
                }
            }
            return (string.Join(',', cl), string.Join(',', vl));
        }
        /// <summary>
        /// 組合 {Column} = {Value} 
        /// </summary>
        /// <param name="param"></param>
        /// <param name="pos">分隔字元或字串</param>
        /// <returns></returns>
        static string ParamsToString(object? param, string pos = ",")
        {
            if (param == null) return string.Empty;
            var rt = new List<string>();
            if (param is DynamicParameters dp)
            {
                foreach (string p in dp.ParameterNames)
                {
                    rt.Add($"{p} = {SetValueString(dp.Get<object?>(p))}");
                }
            }
            else if (param is IDictionary<string, object?> dic)
            {
                foreach (var kvp in dic)
                {
                    rt.Add($"{kvp.Key} = {SetValueString(kvp.Value)}");
                }
            }
            else
            {
                var type = param.GetType();
                foreach (PropertyInfo p in type.GetProperties())
                {
                    rt.Add($"{p.Name} = {SetValueString(p.GetValue(param))}");
                }
            }
            return string.Join(pos, rt);
        }

        static string GetConnectionString(string connstr) => conf.GetConnectionString(connstr) ?? connstr;
        static DbType? GetDbType(object? v)
        {
            if (v == null) return null;
            return v switch
            {
                byte[] or byte => DbType.Binary,
                int or short => DbType.Int32,
                long => DbType.Int64,
                decimal or float => DbType.Decimal,
                double => DbType.Double,
                DateTime or TimeSpan => DbType.DateTime,
                bool => DbType.Boolean,
                _ => DbType.String
            };
        }
    }
}
