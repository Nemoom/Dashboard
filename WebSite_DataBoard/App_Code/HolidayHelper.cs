using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using Newtonsoft.Json;

/// <summary>
///HolidayHelper 的摘要说明
/// </summary>
public class HolidayHelper
{
    #region 字段属性
    private static object _syncObj = new object();
    private static HolidayHelper _instance { get; set; }
    private static List<DateModel> cacheDateList { get; set; }
    
    public HolidayHelper()
	{
		//
		//TODO: 在此处添加构造函数逻辑
		//
	}
    /// <summary>
    /// 获得单例对象,使用懒汉式（双重锁定）
    /// </summary>
    /// <returns></returns>
    public static HolidayHelper GetInstance()
    {
        if (_instance == null)
        {
            lock (_syncObj)
            {
                if (_instance == null)
                {
                    _instance = new HolidayHelper();
                }
            }
        }
        return _instance;
    }
    #endregion
    #region 私有方法
    /// <summary>
    /// 读取文件
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    private string GetFileContent(string filePath)
    {
        string result = "";
        if (File.Exists(filePath))
        {
            result = File.ReadAllText(filePath);
        }
        return result;
    }
    /// <summary>
    /// 获取配置的Json文件
    /// </summary>
    /// <returns>经过反序列化之后的对象集合</returns>
    private List<DateModel> GetConfigList()
    {
        string path = string.Format("{0}Config\\holidayConfig.json", System.AppDomain.CurrentDomain.BaseDirectory);
        string fileContent = GetFileContent(path);
        if (!string.IsNullOrWhiteSpace(fileContent))
        {
            cacheDateList = Newtonsoft.Json.JsonConvert.DeserializeObject<List<DateModel>>(fileContent);
        }
        return cacheDateList;
    }
    /// <summary>
    /// 获取指定年份的数据
    /// </summary>
    /// <param name="year"></param>
    /// <returns></returns>
    private DateModel GetConfigDataByYear(int year)
    {
        if (cacheDateList == null)//取配置数据
            GetConfigList();
        DateModel result = cacheDateList.FirstOrDefault(m => m.Year == year);
        return result;
    }
    /// <summary>
    /// 判断是否为工作日
    /// </summary>
    /// <param name="currDate">要判断的时间</param>
    /// <param name="thisYearData">当前的数据</param>
    /// <returns></returns>
    private bool IsWorkDay(DateTime currDate, DateModel thisYearData)
    {
        if (currDate.Year != thisYearData.Year)//跨年重新读取数据
        {
            thisYearData = GetConfigDataByYear(currDate.Year);
        }
        if (thisYearData.Year > 0)
        {
            string date = currDate.ToString("MMdd");
            int week = (int)currDate.DayOfWeek;

            if (thisYearData.Work.IndexOf(date) >= 0)
            {
                return true;
            }

            if (thisYearData.Holiday.IndexOf(date) >= 0)
            {
                return false;
            }

            if (week != 0 && week != 6)
            {
                return true;
            }
        }
        return false;
    }

    #endregion

    #region Public Method
    public void CleraCacheData()
    {
        if (cacheDateList != null)
        {
            cacheDateList.Clear();
        }
    }
    /// <summary>
    /// 根据传入的工作日天数，获得计算后的日期,可传负数
    /// </summary>
    /// <param name="day">天数</param>
    /// <param name="isContainToday">当天是否算工作日（默认：true）</param>
    /// <returns></returns>
    public DateTime GetReckonDate2Today(int day, bool isContainToday = true)
    {
        DateTime currDate = DateTime.Now;
        int addDay = day >= 0 ? 1 : -1;

        if (isContainToday)
            currDate = currDate.AddDays(-addDay);

        DateModel thisYearData = GetConfigDataByYear(currDate.Year);
        if (thisYearData.Year > 0)
        {
            int sumDay = Math.Abs(day);
            int workDayNum = 0;
            while (workDayNum < sumDay)
            {
                currDate = currDate.AddDays(addDay);
                if (IsWorkDay(currDate, thisYearData))
                    workDayNum++;
            }
        }
        return currDate;
    }
    /// <summary>
    /// 根据传入的日期和工作日天数，获得计算后的日期,可传负数
    /// </summary>
    /// <param name="day">天数</param>
    /// <param name="isContainToday">当天是否算工作日（默认：true）</param>
    /// <returns></returns>
    public DateTime GetReckonDate(DateTime date, int day, bool isContainToday = true)
    {
        if (date==new DateTime(2019,10,24))
        {
            
        }
        DateTime currDate = date;
        int addDay = day >= 0 ? 1 : -1;

        if (isContainToday)
            currDate = currDate.AddDays(-addDay);

        DateModel thisYearData = GetConfigDataByYear(currDate.Year);
        if (thisYearData.Year > 0)
        {
            int sumDay = Math.Abs(day);
            int workDayNum = 0;
            while (workDayNum < sumDay)
            {
                currDate = currDate.AddDays(addDay);
                if (IsWorkDay(currDate, thisYearData))
                    workDayNum++;
            }
        }
        return currDate;
    }
    /// <summary>
    /// 根据传入的时间，计算工作日天数
    /// </summary>
    /// <param name="date">带计算的时间</param>
    /// <param name="isContainToday">当天是否算工作日（默认：true）</param>
    /// <returns></returns>
    public int GetWorkDayNum2Today(DateTime date, bool isContainToday = true)
    {
        var currDate = DateTime.Now;

        int workDayNum = 0;
        int addDay = date.Date > currDate.Date ? 1 : -1;

        if (isContainToday)
        {
            currDate = currDate.AddDays(-addDay);
        }

        DateModel thisYearData = GetConfigDataByYear(currDate.Year);
        if (thisYearData.Year > 0)
        {
            bool isEnd = false;
            do
            {
                currDate = currDate.AddDays(addDay);
                if (IsWorkDay(currDate, thisYearData))
                    workDayNum += addDay;
                isEnd = addDay > 0 ? (date.Date > currDate.Date) : (date.Date < currDate.Date);
            } while (isEnd);
        }
        return workDayNum;
    }
    /// <summary>
    /// 根据传入的两个日期，计算两天间隔的工作日天数
    /// </summary>
    /// <param name="date">带计算的时间</param>
    /// <param name="isContainToday">首日是否算工作日（默认：true）</param>
    /// <returns></returns>
    public int GetWorkDayNum(DateTime date_Start,DateTime date_End, bool isContainToday = true)
    {
        var currDate = date_Start;

        int workDayNum = 0;
        if (date_Start == date_End)
        {
            if (isContainToday)
            { 
                return 1;
            }
            else
            {
                return 0;
            }            
        }
        int addDay = date_End.Date > currDate.Date ? 1 : -1;
        DateModel thisYearData = GetConfigDataByYear(currDate.Year);
        if (isContainToday)
        {
            //currDate = currDate.AddDays(-addDay);
            workDayNum = 1;
            if (!IsWorkDay(currDate, thisYearData))
            { 
                workDayNum = 0; 
            }
        }

        if (thisYearData.Year > 0)
        {
            bool isEnd = false;
            do
            {
                currDate = currDate.AddDays(addDay);
                if (IsWorkDay(currDate, thisYearData))
                    workDayNum += addDay;
                isEnd = addDay > 0 ? (date_End.Date > currDate.Date) : (date_End.Date < currDate.Date);
            } while (isEnd);
        }
        return workDayNum;
    }
    #endregion
}

public struct DateModel
{
    public int Year { get; set; }

    public List<string> Work { get; set; }

    public List<string> Holiday { get; set; }
}