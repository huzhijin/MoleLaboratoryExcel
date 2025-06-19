using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;

public class LogDao
{
    public List<SystemLog> GetLogs(DateTime? startTime = null, DateTime? endTime = null, 
        string username = null, string action = null)
    {
        var sql = "SELECT * FROM SystemLogs WHERE 1=1";
        var parameters = new List<SqlParameter>();

        if (startTime.HasValue)
        {
            sql += " AND LogTime >= @StartTime";
            parameters.Add(new SqlParameter("@StartTime", startTime.Value));
        }

        if (endTime.HasValue)
        {
            sql += " AND LogTime <= @EndTime";
            parameters.Add(new SqlParameter("@EndTime", endTime.Value));
        }

        if (!string.IsNullOrEmpty(username))
        {
            sql += " AND Username LIKE @Username";
            parameters.Add(new SqlParameter("@Username", $"%{username}%"));
        }

        if (!string.IsNullOrEmpty(action))
        {
            sql += " AND Action LIKE @Action";
            parameters.Add(new SqlParameter("@Action", $"%{action}%"));
        }

        sql += " ORDER BY LogTime DESC";

        var logs = new List<SystemLog>();
        try
        {
            var dt = DbHelper.ExecuteQuery(sql, parameters.ToArray());
            foreach (DataRow row in dt.Rows)
            {
                logs.Add(MapDataRowToLog(row));
            }
        }
        catch (Exception ex)
        {
            LogHelper.LogError("获取日志列表失败", ex);
            throw;
        }
        return logs;
    }
    private string ConvertActionToChinese(string action)
    {
        switch (action)
        {
            case "Login": return "登录";
            case "AddUser": return "新增用户";
            case "UpdateUser": return "修改用户";
            case "DeleteUser": return "删除用户";
            case "ImportExcel": return "导入Excel";
            case "ExportWord": return "导出Word";
            case "Error": return "错误";
            default: return action;
        }
    }


    private SystemLog MapDataRowToLog(DataRow row)
    {
        return new SystemLog
        {
            Id = Convert.ToInt32(row["Id"]),
            UserId = row["UserId"] != DBNull.Value ? Convert.ToInt32(row["UserId"]) : (int?)null,
            Username = row["Username"]?.ToString(),
            Action = row["Action"].ToString(),
            Description = row["Description"].ToString(),
            LogTime = Convert.ToDateTime(row["LogTime"]),
            IPAddress = row["IPAddress"]?.ToString()
        };
    }
} 