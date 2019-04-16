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
    //ECharts4:Finish POs
    public static int[] count_PO_Finished_perMonth = new int[12];
    public static double[] NetValue_PO_Finished_perMonth = new double[12];
    public static int[] count_PO_ForecastFinished_perMonth = new int[12];
    //ECharts5:SO Create perMonth
    public static int[] count_SO_soCreated_perMonth = new int[12];
    public static double[] NetValue_SO_soCreated_perMonth = new double[12];
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
   
}