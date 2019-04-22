using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
///Data2Trace 的摘要说明
/// </summary>
public class Data2Trace
{
	public Data2Trace()
	{
		//
		//TODO: 在此处添加构造函数逻辑
		//        
	}
    public static string str_PO_soCreated_perMonth 
    {
        get 
        {
            string mstr_PO_soCreated_perMonth = "";
            for (int i = 0; i < count_PO_soCreated_perMonth.Length; i++)
            {
                mstr_PO_soCreated_perMonth = mstr_PO_soCreated_perMonth + count_PO_soCreated_perMonth[i].ToString() + "!";
            }
            return mstr_PO_soCreated_perMonth.Substring(0, mstr_PO_soCreated_perMonth.Length - 1); 
        } 
    }
    public static string str_CreatedDelay_perMonth
    {
        get
        {
            string mstr_CreatedDelay_perMonth = "";
            for (int i = 0; i < count_CreatedDelay_perMonth.Length; i++)
            {
                mstr_CreatedDelay_perMonth = mstr_CreatedDelay_perMonth + count_CreatedDelay_perMonth[i].ToString() + "!";
            }
            return mstr_CreatedDelay_perMonth.Substring(0, mstr_CreatedDelay_perMonth.Length - 1);
        }
    }
    public static string str_PO_poCreated_perMonth
    {
        get
        {
            string mstr_PO_poCreated_perMonth = "";
            for (int i = 0; i < count_PO_poCreated_perMonth.Length; i++)
            {
                mstr_PO_poCreated_perMonth = mstr_PO_poCreated_perMonth + count_PO_poCreated_perMonth[i].ToString() + "!";
            }
            return mstr_PO_poCreated_perMonth.Substring(0, mstr_PO_poCreated_perMonth.Length - 1);
        }
    }
    public static string str_ReleaseDelay_perMonth
    {
        get
        {
            string mstr_ReleaseDelay_perMonth = "";
            for (int i = 0; i < count_ReleaseDelay_perMonth.Length; i++)
            {
                mstr_ReleaseDelay_perMonth = mstr_ReleaseDelay_perMonth + count_ReleaseDelay_perMonth[i].ToString() + "!";
            }
            return mstr_ReleaseDelay_perMonth.Substring(0, mstr_ReleaseDelay_perMonth.Length - 1);
        }
    }
    public static string str_DCR_inside
    {
        get
        {
            return count_PO_DIV.ToString() + "!" + count_PO_noDIV.ToString();
        }
    }
    public static string str_DCR_outside
    {
        get
        {
            return count_PO_DIV_OnTime.ToString()
                + "!" + count_PO_DIV_Delay.ToString()
                + "!" + count_PO_noDIV_Delay.ToString()
                + "!" + count_PO_noDIV_Warning.ToString()
                + "!" + count_PO_noDIV_OK.ToString();
        }
    }
    public static string str_PO_Finished_perMonth
    {
        get
        {
            string mstr_PO_Finished_perMonth = "";
            for (int i = 0; i < count_PO_Finished_perMonth.Length; i++)
            {
                mstr_PO_Finished_perMonth = mstr_PO_Finished_perMonth + count_PO_Finished_perMonth[i].ToString() + "!";
            }
            return mstr_PO_Finished_perMonth.Substring(0, mstr_PO_Finished_perMonth.Length - 1);
        }
    }
    public static string str_PO_ForecastFinished_perMonth
    {
        get
        {
            string mstr_PO_ForecastFinished_perMonth = "";
            for (int i = 0; i < count_PO_ForecastFinished_perMonth.Length; i++)
            {
                mstr_PO_ForecastFinished_perMonth = mstr_PO_ForecastFinished_perMonth + count_PO_ForecastFinished_perMonth[i].ToString() + "!";
            }
            return mstr_PO_ForecastFinished_perMonth.Substring(0, mstr_PO_ForecastFinished_perMonth.Length - 1);
        }
    }
    public static string str_NetValue_PO_Finished_perMonth
    {
        get
        {
            string mstr_NetValue_PO_Finished_perMonth = "";
            for (int i = 0; i < NetValue_PO_Finished_perMonth.Length; i++)
            {
                mstr_NetValue_PO_Finished_perMonth = mstr_NetValue_PO_Finished_perMonth + NetValue_PO_Finished_perMonth[i].ToString("F2") + "!";
            }
            return mstr_NetValue_PO_Finished_perMonth.Substring(0, mstr_NetValue_PO_Finished_perMonth.Length - 1);
        }
    }
    public static string str_SO_soCreated_perMonth
    {
        get
        {
            string mstr_SO_soCreated_perMonth = "";
            for (int i = 0; i < count_SO_soCreated_perMonth.Length; i++)
            {
                mstr_SO_soCreated_perMonth = mstr_SO_soCreated_perMonth + count_SO_soCreated_perMonth[i].ToString() + "!";
            }
            return mstr_SO_soCreated_perMonth.Substring(0, mstr_SO_soCreated_perMonth.Length - 1);
        }
    }
    public static string str_NetValue_SO_soCreated_perMonth
    {
        get
        {
            string mstr_NetValue_SO_soCreated_perMonth = "";
            for (int i = 0; i < NetValue_SO_soCreated_perMonth.Length; i++)
            {
                mstr_NetValue_SO_soCreated_perMonth = mstr_NetValue_SO_soCreated_perMonth + NetValue_SO_soCreated_perMonth[i].ToString("F2") + "!";
            }
            return mstr_NetValue_SO_soCreated_perMonth.Substring(0, mstr_NetValue_SO_soCreated_perMonth.Length - 1);
        }
    }
    public static string str_LT_perMonth
    {
        get
        {
            string mstr_LT_perMonth = "";
            for (int i = 0; i < LT_perMonth.Length; i++)
            {
                mstr_LT_perMonth = mstr_LT_perMonth + LT_perMonth[i].ToString("F2") + "!";
            }
            return mstr_LT_perMonth.Substring(0, mstr_LT_perMonth.Length - 1);
        }
    }
    public static string str_LT_perWeek
    {
        get
        {
            string mstr_LT_perWeek = "";
            for (int i = 0; i < LT_perWeek.Length; i++)
            {
                mstr_LT_perWeek = mstr_LT_perWeek + LT_perWeek[i].ToString("F2") + "!";
            }
            return mstr_LT_perWeek.Substring(0, mstr_LT_perWeek.Length - 1);
        }
    }

    public static string values_ECharts1
    {
        get
        {
            return str_PO_soCreated_perMonth + ";" + str_CreatedDelay_perMonth + ";" + (count_PO_soCreated_YTD - count_CreatedDelay_YTD).ToString() + "!" + count_CreatedDelay_YTD.ToString();
        }
    }
    public static string values_ECharts2
    {
        get
        {
            return str_PO_poCreated_perMonth + ";" + str_ReleaseDelay_perMonth + ";" +(count_PO_poCreated_YTD - count_ReleaseDelay_YTD).ToString() + "!" +count_ReleaseDelay_YTD.ToString();
        }
    }
    public static string values_ECharts3
    {
        get
        {
            return str_DCR_inside + ";" + str_DCR_outside;
        }
    }
    public static string values_ECharts4
    {
        get
        {
            return str_PO_Finished_perMonth + "!" + count_PO_Finished_YTD.ToString()+";"
                + str_PO_ForecastFinished_perMonth + "!" + count_PO_ForecastFinished_YTD.ToString() + ";"
                + str_NetValue_PO_Finished_perMonth + "!" + NetValue_PO_Finished_YTD.ToString("F2");
        }
    }
    public static string values_ECharts5
    {
        get
        {
            return str_SO_soCreated_perMonth + "!" + count_SO_soCreated_YTD.ToString() + ";"
                + str_NetValue_SO_soCreated_perMonth + "!" + NetValue_SO_soCreated_YTD.ToString("F2");
        }
    }
    public static string values_ECharts6
    {
        get
        {
            return str_LT_perMonth + "!" + str_LT_perWeek + "!" + LT_YTD.ToString("F2");
        }
    }
    public static string values_ECharts7
    {
        get
        {
            return str_DCR_inside + ";" + str_DCR_outside;
        }
    }
    public static string values_ECharts8
    {
        get
        {
            return count_Stored_5 + "!" + count_Stored_5_10 + "!" + count_Stored_10_20 + "!" + count_Stored_20 + "!" + count_Stored;
        }
    }
    //ECharts1:PO Create Monitor
    public static int[] count_PO_soCreated_perMonth = new int[12];
    public static int[] count_CreatedDelay_perMonth = new int[12];
    public static int count_PO_soCreated_YTD;
    public static int count_CreatedDelay_YTD;
    //ECharts2:PO Release Monitor
    public static int[] count_PO_poCreated_perMonth = new int[12];
    public static int[] count_ReleaseDelay_perMonth = new int[12];
    public static int count_PO_poCreated_YTD;
    public static int count_ReleaseDelay_YTD;
    //ECharts3:DCR(Delivery Class Reliablity)
    public static int count_PO_DIV;
    public static int count_PO_DIV_OnTime;
    public static int count_PO_DIV_Delay;
    public static int count_PO_noDIV;
    public static int count_PO_noDIV_Delay;
    public static int count_PO_noDIV_Warning;
    public static int count_PO_noDIV_OK;
    public static int count_PO_LastFailed;//昨天刚刚超期的订单数
    //ECharts4:Finish POs
    public static int[] count_PO_Finished_perMonth = new int[12];
    public static int count_PO_Finished_YTD;
    public static double[] NetValue_PO_Finished_perMonth = new double[12];
    public static double NetValue_PO_Finished_YTD;
    public static int[] count_PO_ForecastFinished_perMonth = new int[12];
    public static int count_PO_ForecastFinished_YTD;
    //ECharts5:SO Create perMonth
    public static int[] count_SO_soCreated_perMonth = new int[12];//需要查重，从List<string>中处理一下
    public static int count_SO_soCreated_YTD;//需要查重，从List<string>中处理一下
    public static double[] NetValue_SO_soCreated_perMonth = new double[12];
    public static double NetValue_SO_soCreated_YTD;
    //ECharts6:LT(Lead Time)
    public static double[] LT_perMonth = new double[12];
    public static double[] LT_perWeek = new double[4];
    public static double LT_YTD;
    //ECharts7:Reminders for 3Weeks&2Weeks
    public static int count_Reminder3Weeks;
    public static int count_Reminder2Weeks;
    //ECharts8:Repertory Monitor
    public static int count_Stored_5;
    public static int count_Stored_5_10;
    public static int count_Stored_10_20;
    public static int count_Stored_20;
    public static int count_Stored;//已完成无ShipmentStartOn

    public void CleanCounter()
    {
        //ECharts1:PO Create Monitor
        count_PO_soCreated_perMonth = new int[12];
        count_CreatedDelay_perMonth = new int[12];
        count_PO_soCreated_YTD = 0;
        count_CreatedDelay_YTD = 0;
        //ECharts2:PO Release Monitor
        count_PO_poCreated_perMonth = new int[12];
        count_ReleaseDelay_perMonth = new int[12];
        count_PO_poCreated_YTD = 0;
        count_ReleaseDelay_YTD = 0;
        //ECharts3:DCR(Delivery Class Reliablity)
        count_PO_DIV = 0;
        count_PO_DIV_OnTime = 0;
        count_PO_DIV_Delay = 0;
        count_PO_noDIV = 0;
        count_PO_noDIV_Delay = 0;
        count_PO_noDIV_Warning = 0;
        count_PO_noDIV_OK = 0;
        count_PO_LastFailed = 0;//昨天刚刚超期的订单数
        //ECharts4:Finish POs
        count_PO_Finished_perMonth = new int[12];
        NetValue_PO_Finished_perMonth = new double[12];
        count_PO_ForecastFinished_perMonth = new int[12];
        count_PO_Finished_YTD = 0;
        NetValue_PO_Finished_YTD = 0;
        count_PO_ForecastFinished_YTD = 0;
        //ECharts5:SO Create perMonth
        count_SO_soCreated_perMonth = new int[12];
        NetValue_SO_soCreated_perMonth = new double[12];
        count_SO_soCreated_YTD = 0;
        NetValue_SO_soCreated_YTD = 0;
        //ECharts6:LT(Lead Time)
        LT_perMonth = new double[12];
        LT_perWeek = new double[4];
        LT_YTD = 0;
        //ECharts7:Reminders for 3Weeks&2Weeks
        count_Reminder3Weeks = 0;
        count_Reminder2Weeks = 0;
        //ECharts8:Repertory Monitor
        count_Stored_5 = 0;
        count_Stored_5_10 = 0;
        count_Stored_10_20 = 0;
        count_Stored_20 = 0;
        count_Stored = 0;
    }
}