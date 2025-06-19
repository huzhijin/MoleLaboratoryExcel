using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;

public class UserDao
{
    public User GetUserByUsername(string username)
    {
        var sql = "SELECT * FROM Users WHERE Username = @Username";
        var parameters = new[] { new SqlParameter("@Username", username) };

        try
        {
            var dt = DbHelper.ExecuteQuery(sql, parameters);
            if (dt.Rows.Count > 0)
            {
                return MapDataRowToUser(dt.Rows[0]);
            }
        }
        catch (Exception ex)
        {
            LogHelper.LogError($"获取用户信息失败: {username}", ex);
            throw;
        }
        return null;
    }

    public List<User> GetAllUsers()
    {
        var sql = "SELECT * FROM Users ORDER BY CreateTime DESC";
        var users = new List<User>();

        try
        {
            var dt = DbHelper.ExecuteQuery(sql);
            foreach (DataRow row in dt.Rows)
            {
                users.Add(MapDataRowToUser(row));
            }
        }
        catch (Exception ex)
        {
            LogHelper.LogError("获取用户列表失败", ex);
            throw;
        }
        return users;
    }

    public bool AddUser(User user)
    {
        var sql = @"INSERT INTO Users (Username, Password, Role, CreateTime, IsActive) 
                   VALUES (@Username, @Password, @Role, @CreateTime, @IsActive)";

        var parameters = new[]
        {
            new SqlParameter("@Username", user.Username),
            new SqlParameter("@Password", user.Password),
            new SqlParameter("@Role", user.Role),
            new SqlParameter("@CreateTime", DateTime.Now),
            new SqlParameter("@IsActive", user.IsActive)
        };

        try
        {
            return DbHelper.ExecuteNonQuery(sql, parameters) > 0;
        }
        catch (Exception ex)
        {
            LogHelper.LogError($"添加用户失败: {user.Username}", ex);
            throw;
        }
    }

    public bool UpdateUser(User user)
    {
        var sql = @"UPDATE Users SET 
                    Username = @Username,
                    Role = @Role,
                    IsActive = @IsActive
                    WHERE Id = @Id";

        if (!string.IsNullOrEmpty(user.Password))
        {
            sql = sql.Replace("WHERE", ", Password = @Password WHERE");
        }

        var parameters = new List<SqlParameter>
        {
            new SqlParameter("@Id", user.Id),
            new SqlParameter("@Username", user.Username),
            new SqlParameter("@Role", user.Role),
            new SqlParameter("@IsActive", user.IsActive)
        };

        if (!string.IsNullOrEmpty(user.Password))
        {
            parameters.Add(new SqlParameter("@Password", user.Password));
        }

        try
        {
            return DbHelper.ExecuteNonQuery(sql, parameters.ToArray()) > 0;
        }
        catch (Exception ex)
        {
            LogHelper.LogError($"更新用户失败: {user.Username}", ex);
            throw;
        }
    }

    public bool DeleteUser(int userId)
    {
        var sql = "DELETE FROM Users WHERE Id = @Id";
        var parameters = new[] { new SqlParameter("@Id", userId) };

        try
        {
            return DbHelper.ExecuteNonQuery(sql, parameters) > 0;
        }
        catch (Exception ex)
        {
            LogHelper.LogError($"删除用户失败: ID={userId}", ex);
            throw;
        }
    }

    private User MapDataRowToUser(DataRow row)
    {
        return new User
        {
            Id = Convert.ToInt32(row["Id"]),
            Username = row["Username"].ToString(),
            Password = row["Password"].ToString(),
            Role = row["Role"].ToString(),
            CreateTime = Convert.ToDateTime(row["CreateTime"]),
            IsActive = Convert.ToBoolean(row["IsActive"])
        };
    }
} 