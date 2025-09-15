using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;

public static class LogHelper
{
    public static void LogUserAction(string username, string action, string description)
    {
        // Ensure using English action types
        string englishAction = ConvertActionToEnglish(action);

        try
        {
            var sql = @"INSERT INTO SystemLogs (Username, Action, Description, LogTime, IPAddress) 
                       VALUES (@Username, @Action, @Description, @LogTime, @IPAddress)";

            var parameters = new[]
            {
                new SqlParameter("@Username", username),
                new SqlParameter("@Action", englishAction),
                new SqlParameter("@Description", description),
                new SqlParameter("@LogTime", DateTime.Now),
                new SqlParameter("@IPAddress", GetClientIP())
            };

            DbHelper.ExecuteNonQuery(sql, parameters);
        }
        catch (Exception ex)
        {
            // Log error to file
            LogError("Record user operation failed", ex);
        }
    }

    private static string ConvertActionToEnglish(string action)
    {
        switch (action)
        {
            case "登录":
            case "用户登录成功":
            case "管理员登录成功":
                return "Login";
            case "新增用户":
            case "添加用户":
                return "AddUser";
            case "修改用户":
                return "UpdateUser";
            case "删除用户":
                return "DeleteUser";
            case "导入Excel":
            case "导入Excel数据":
                return "ImportExcel";
            case "导出Word":
            case "生成Word报告":
                return "ExportWord";
            case "错误":
            case "登录失败":
                return "Error";
            case "批量重命名":
            case "BatchRename":
                return "BatchRename";
            default:
                return action;
        }
    }

    public static void LogError(string message, Exception ex)
    {
        try
        {
            var sql = @"INSERT INTO SystemLogs (Username, Action, Description, LogTime, IPAddress) 
                       VALUES (@Username, @Action, @Description, @LogTime, @IPAddress)";

            var description = $"{message}: {ex.Message}";
            if (ex.InnerException != null)
            {
                description += $" | Inner Exception: {ex.InnerException.Message}";
            }

            var parameters = new[]
            {
                new SqlParameter("@Username", "System"),
                new SqlParameter("@Action", "Error"),
                new SqlParameter("@Description", description),
                new SqlParameter("@LogTime", DateTime.Now),
                new SqlParameter("@IPAddress", GetClientIP())
            };

            DbHelper.ExecuteNonQuery(sql, parameters);
        }
        catch
        {
            // If database record fails, write to file log
            var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            var logFile = Path.Combine(logPath, $"Error_{DateTime.Now:yyyyMMdd}.log");
            var logContent = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\r\n{ex}\r\n\r\n";

            File.AppendAllText(logFile, logContent);
        }
    }

    private static string GetClientIP()
    {
        try
        {
            var host = Dns.GetHostEntry(Dns.GetHostName());
            return host.AddressList.FirstOrDefault(ip => ip.AddressFamily == AddressFamily.InterNetwork)?.ToString() ?? "Unknown";
        }
        catch
        {
            return "Unknown";
        }
    }
}