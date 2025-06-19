using System;

public class SystemLog
{
    public int Id { get; set; }
    public int? UserId { get; set; }
    public string Username { get; set; }
    public string Action { get; set; }
    public string Description { get; set; }
    public DateTime LogTime { get; set; }
    public string IPAddress { get; set; }
} 