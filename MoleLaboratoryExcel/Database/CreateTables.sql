-- 用户表
CREATE TABLE [dbo].[Users](
    [Id] [int] IDENTITY(1,1) PRIMARY KEY,
    [Username] [nvarchar](50) NOT NULL,
    [Password] [nvarchar](256) NOT NULL,
    [Role] [nvarchar](20) NOT NULL,
    [CreateTime] [datetime] NOT NULL,
    [IsActive] [bit] NOT NULL,
    [LastLoginTime] [datetime] NULL
)

-- 日志表
CREATE TABLE [dbo].[SystemLogs](
    [Id] [int] IDENTITY(1,1) PRIMARY KEY,
    [UserId] [int] NULL,
    [Username] [nvarchar](50) NULL,
    [Action] [nvarchar](50) NOT NULL,
    [Description] [nvarchar](500) NULL,
    [LogTime] [datetime] NOT NULL,
    [IPAddress] [nvarchar](50) NULL
) 