using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Windows.Forms;

public static class DbHelper
{
    private static string _connectionString;

    static DbHelper()
    {
        try
        {
            // 首先尝试从配置文件读取连接字符串
            _connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"]?.ConnectionString;

            
        }
        catch (Exception ex)
        {
            MessageBox.Show("读取数据库配置失败：" + ex.Message, "错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            LogHelper.LogError("读取数据库配置失败", ex);
        }
    }

    // 提供修改连接字符串的方法
    public static void UpdateConnectionString(string server, string database, string username, string password)
    {
        _connectionString = $"Data Source={server};Initial Catalog={database};User ID={username};Password={password};Connect Timeout=30;";
    }

    public static SqlConnection GetConnection()
    {
        return new SqlConnection(_connectionString);
    }

    public static DataTable ExecuteQuery(string sql, SqlParameter[] parameters = null)
    {
        using (var connection = GetConnection())
        {
            try
            {
                connection.Open();
                using (var command = new SqlCommand(sql, connection))
                {
                    if (parameters != null)
                    {
                        command.Parameters.AddRange(parameters);
                    }

                    var dataTable = new DataTable();
                    using (var adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                    return dataTable;
                }
            }
            catch (Exception ex)
            {
                LogHelper.LogError($"执行查询失败: {sql}", ex);
                throw;
            }
        }
    }

    public static int ExecuteNonQuery(string sql, SqlParameter[] parameters = null)
    {
        using (var connection = GetConnection())
        {
            try
            {
                connection.Open();
                using (var command = new SqlCommand(sql, connection))
                {
                    if (parameters != null)
                    {
                        command.Parameters.AddRange(parameters);
                    }
                    return command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                LogHelper.LogError($"执行非查询操作失败: {sql}", ex);
                throw;
            }
        }
    }

    public static object ExecuteScalar(string sql, SqlParameter[] parameters = null)
    {
        using (var connection = GetConnection())
        {
            try
            {
                connection.Open();
                using (var command = new SqlCommand(sql, connection))
                {
                    if (parameters != null)
                    {
                        command.Parameters.AddRange(parameters);
                    }
                    return command.ExecuteScalar();
                }
            }
            catch (Exception ex)
            {
                LogHelper.LogError($"执行标量查询失败: {sql}", ex);
                throw;
            }
        }
    }

    // 测试数据库连接
    public static bool TestConnection(string connectionString = null)
    {
        using (var connection = new SqlConnection(connectionString ?? _connectionString))
        {
            try
            {
                connection.Open();
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.LogError("测试数据库连接失败", ex);
                return false;
            }
        }
    }
}