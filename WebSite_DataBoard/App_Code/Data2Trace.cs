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
    #region Array String with "!"
    public static string str_PO_soCreated_perMonth
    {
        get
        {
            string mstr_PO_soCreated_perMonth = "";
            for (int i = 0; i < count_PO_soCreated_0400_perMonth.Length; i++)
            {
                mstr_PO_soCreated_perMonth = mstr_PO_soCreated_perMonth + count_PO_soCreated_0400_perMonth[i].ToString() + "!";
            }
            return mstr_PO_soCreated_perMonth.Substring(0, mstr_PO_soCreated_perMonth.Length - 1);
        }
    }
    public static string str_CreatedDelay_perMonth
    {
        get
        {
            string mstr_CreatedDelay_perMonth = "";
            for (int i = 0; i < count_CreatedDelay_0400_perMonth.Length; i++)
            {
                mstr_CreatedDelay_perMonth = mstr_CreatedDelay_perMonth + count_CreatedDelay_0400_perMonth[i].ToString() + "!";
            }
            return mstr_CreatedDelay_perMonth.Substring(0, mstr_CreatedDelay_perMonth.Length - 1);
        }
    }
    public static string str_PO_poCreated_perMonth
    {
        get
        {
            string mstr_PO_poCreated_perMonth = "";
            for (int i = 0; i < count_PO_poCreated_0400_perMonth.Length; i++)
            {
                mstr_PO_poCreated_perMonth = mstr_PO_poCreated_perMonth + count_PO_poCreated_0400_perMonth[i].ToString() + "!";
            }
            return mstr_PO_poCreated_perMonth.Substring(0, mstr_PO_poCreated_perMonth.Length - 1);
        }
    }
    public static string str_ReleaseDelay_perMonth
    {
        get
        {
            string mstr_ReleaseDelay_perMonth = "";
            for (int i = 0; i < count_ReleaseDelay_0400_perMonth.Length; i++)
            {
                mstr_ReleaseDelay_perMonth = mstr_ReleaseDelay_perMonth + count_ReleaseDelay_0400_perMonth[i].ToString() + "!";
            }
            return mstr_ReleaseDelay_perMonth.Substring(0, mstr_ReleaseDelay_perMonth.Length - 1);
        }
    }
    //public static string str_DCR_inside
    //{
    //    get
    //    {
    //        return count_PO_DLV.ToString() + "!" + count_PO_noDLV.ToString();
    //    }
    //}
    //public static string str_DCR_outside
    //{
    //    get
    //    {
    //        return count_PO_DLV_OnTime.ToString()
    //            + "!" + count_PO_DLV_Delay.ToString()
    //            + "!" + count_PO_noDLV_Delay.ToString()
    //            + "!" + count_PO_noDLV_Warning.ToString()
    //            + "!" + count_PO_noDLV_OK.ToString();
    //    }
    //}
    public static string str_DCR_0400
    {
        get
        {
            string mstr_DCR_0400 = "";
            for (int i = 0; i < count_0400_PO_Finished_perMonth.Length; i++)
            {
                if (count_0400_PO_Finished_perMonth[i] == 0)
                {
                    //分母为0
                    mstr_DCR_0400 = mstr_DCR_0400 + "0!";
                }
                else
                {
                    //mstr_DCR_0400 = mstr_DCR_0400 + Convert.ToInt16(((Convert.ToDouble(count_0400_PO_OnTime_perMonth[i]) / 
                    //    Convert.ToDouble(count_0400_PO_Finished_perMonth[i]))*100)).ToString("") + "!";
                    mstr_DCR_0400 = mstr_DCR_0400 + ((int)((Convert.ToDouble(count_0400_PO_OnTime_perMonth[i]) /
                        Convert.ToDouble(count_0400_PO_Finished_perMonth[i])) * 100 + 0.5)).ToString("") + "!";
                }                
            }
            return mstr_DCR_0400.Substring(0, mstr_DCR_0400.Length - 1);
        }
    }
    public static string str_DCR_0481
    {
        get
        {
            string mstr_DCR_0481 = "";
            for (int i = 0; i < count_0481_PO_Finished_perMonth.Length; i++)
            {
                if (count_0481_PO_Finished_perMonth[i] == 0)
                {
                    //分母为0
                    mstr_DCR_0481 = mstr_DCR_0481 + "0!";
                }
                else
                {
                    //mstr_DCR_0481 = mstr_DCR_0481 + Convert.ToInt16(((Convert.ToDouble(count_0481_PO_OnTime_perMonth[i]) /
                    //    Convert.ToDouble(count_0481_PO_Finished_perMonth[i])) * 100)).ToString() + "!";
                    mstr_DCR_0481 = mstr_DCR_0481 + ((int)((Convert.ToDouble(count_0481_PO_OnTime_perMonth[i]) /
                        Convert.ToDouble(count_0481_PO_Finished_perMonth[i])) * 100 + 0.5)).ToString() + "!";
                }
            }
            return mstr_DCR_0481.Substring(0, mstr_DCR_0481.Length - 1);
        }
    }
    public static string str_DCR
    {
        get
        {
            string mstr_DCR = "";
            for (int i = 0; i < count_PO_Finished_perMonth.Length; i++)
            {
                if (count_PO_Finished_perMonth[i] == 0)
                {
                    //分母为0
                    mstr_DCR = mstr_DCR + "0!";
                }
                else
                {
                    //mstr_DCR = mstr_DCR + Convert.ToInt16(((Convert.ToDouble(count_0400_PO_OnTime_perMonth[i] +
                    //    count_0481_PO_OnTime_perMonth[i]) / Convert.ToDouble(count_PO_Finished_perMonth[i])) * 100)).ToString() + "!";
                    mstr_DCR = mstr_DCR + ((int)((Convert.ToDouble(count_0400_PO_OnTime_perMonth[i] +
                        count_0481_PO_OnTime_perMonth[i]) / Convert.ToDouble(count_PO_Finished_perMonth[i])) * 100 + 0.5)).ToString() + "!";
                }
            }
            return mstr_DCR.Substring(0, mstr_DCR.Length - 1);
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
    public static string str_PO_Finished_0400_perMonth
    {
        get
        {
            string mstr_PO_Finished_0400_perMonth = "";
            for (int i = 0; i < count_PO_Finished_0400_perMonth.Length; i++)
            {
                mstr_PO_Finished_0400_perMonth = mstr_PO_Finished_0400_perMonth + count_PO_Finished_0400_perMonth[i].ToString() + "!";
            }
            return mstr_PO_Finished_0400_perMonth.Substring(0, mstr_PO_Finished_0400_perMonth.Length - 1);
        }
    }
    public static string str_PO_Finished_0481_perMonth
    {
        get
        {
            string mstr_PO_Finished_0481_perMonth = "";
            for (int i = 0; i < count_PO_Finished_0481_perMonth.Length; i++)
            {
                mstr_PO_Finished_0481_perMonth = mstr_PO_Finished_0481_perMonth + count_PO_Finished_0481_perMonth[i].ToString() + "!";
            }
            return mstr_PO_Finished_0481_perMonth.Substring(0, mstr_PO_Finished_0481_perMonth.Length - 1);
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
    public static string str_SO_soCreated_0400_perMonth
    {
        get
        {
            string mstr_SO_soCreated_0400_perMonth = "";
            for (int i = 0; i < count_SO_soCreated_0400_perMonth.Length; i++)
            {
                mstr_SO_soCreated_0400_perMonth = mstr_SO_soCreated_0400_perMonth + count_SO_soCreated_0400_perMonth[i].ToString() + "!";
            }
            return mstr_SO_soCreated_0400_perMonth.Substring(0, mstr_SO_soCreated_0400_perMonth.Length - 1);
        }
    }
    public static string str_SO_soCreated_0481_perMonth
    {
        get
        {
            string mstr_SO_soCreated_0481_perMonth = "";
            for (int i = 0; i < count_SO_soCreated_0481_perMonth.Length; i++)
            {
                mstr_SO_soCreated_0481_perMonth = mstr_SO_soCreated_0481_perMonth + count_SO_soCreated_0481_perMonth[i].ToString() + "!";
            }
            return mstr_SO_soCreated_0481_perMonth.Substring(0, mstr_SO_soCreated_0481_perMonth.Length - 1);
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
                mstr_LT_perMonth = mstr_LT_perMonth + Convert.ToInt16(LT_perMonth[i]).ToString() + "!";
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
                mstr_LT_perWeek = mstr_LT_perWeek + Convert.ToInt16(LT_perWeek[i]).ToString() + "!";
            }
            return mstr_LT_perWeek.Substring(0, mstr_LT_perWeek.Length - 1);
        }
    }
    public static string str_LT_0400_perMonth
    {
        get
        {
            string mstr_LT_0400_perMonth = "";
            for (int i = 0; i < LT_0400_perMonth.Length; i++)
            {
                mstr_LT_0400_perMonth = mstr_LT_0400_perMonth + Convert.ToInt16(LT_0400_perMonth[i]).ToString() + "!";
            }
            return mstr_LT_0400_perMonth.Substring(0, mstr_LT_0400_perMonth.Length - 1);
        }
    }
    public static string str_LT_0400_perWeek
    {
        get
        {
            string mstr_LT_0400_perWeek = "";
            for (int i = 0; i < LT_0400_perWeek.Length; i++)
            {
                mstr_LT_0400_perWeek = mstr_LT_0400_perWeek + Convert.ToInt16(LT_0400_perWeek[i]).ToString() + "!";
            }
            return mstr_LT_0400_perWeek.Substring(0, mstr_LT_0400_perWeek.Length - 1);
        }
    }
    public static string str_LT_0481_perMonth
    {
        get
        {
            string mstr_LT_0481_perMonth = "";
            for (int i = 0; i < LT_0481_perMonth.Length; i++)
            {
                mstr_LT_0481_perMonth = mstr_LT_0481_perMonth + Convert.ToInt16(LT_0481_perMonth[i]).ToString() + "!";
            }
            return mstr_LT_0481_perMonth.Substring(0, mstr_LT_0481_perMonth.Length - 1);
        }
    }
    public static string str_LT_0481_perWeek
    {
        get
        {
            string mstr_LT_0481_perWeek = "";
            for (int i = 0; i < LT_0481_perWeek.Length; i++)
            {
                mstr_LT_0481_perWeek = mstr_LT_0481_perWeek + Convert.ToInt16(LT_0481_perWeek[i]).ToString() + "!";
            }
            return mstr_LT_0481_perWeek.Substring(0, mstr_LT_0481_perWeek.Length - 1);
        }
    }
    public static string str_averageTime_Ex_plant_0400_perMonth
    {
        get
        {
            string mstr_averageTime_Ex_plant_0400_perMonth = "";
            for (int i = 0; i < averageTime_Ex_plant_0400_perMonth.Length; i++)
            {
                mstr_averageTime_Ex_plant_0400_perMonth = mstr_averageTime_Ex_plant_0400_perMonth + Convert.ToInt16(averageTime_Ex_plant_0400_perMonth[i]).ToString() + "!";
            }
            return mstr_averageTime_Ex_plant_0400_perMonth.Substring(0, mstr_averageTime_Ex_plant_0400_perMonth.Length - 1);
        }
    }
    public static string str_averageTime_Ex_plant_0481_perMonth
    {
        get
        {
            string mstr_averageTime_Ex_plant_0481_perMonth = "";
            for (int i = 0; i < averageTime_Ex_plant_0481_perMonth.Length; i++)
            {
                mstr_averageTime_Ex_plant_0481_perMonth = mstr_averageTime_Ex_plant_0481_perMonth + Convert.ToInt16(averageTime_Ex_plant_0481_perMonth[i]).ToString() + "!";
            }
            return mstr_averageTime_Ex_plant_0481_perMonth.Substring(0, mstr_averageTime_Ex_plant_0481_perMonth.Length - 1);
        }
    }
    public static string str_averageTime_Ex_plant_perMonth
    {
        get
        {
            string mstr_averageTime_Ex_plant_perMonth = "";
            for (int i = 0; i < averageTime_Ex_plant_perMonth.Length; i++)
            {
                mstr_averageTime_Ex_plant_perMonth = mstr_averageTime_Ex_plant_perMonth + Convert.ToInt16(averageTime_Ex_plant_perMonth[i]).ToString() + "!";
            }
            return mstr_averageTime_Ex_plant_perMonth.Substring(0, mstr_averageTime_Ex_plant_perMonth.Length - 1);
        }
    }

    public static string str_averageTime_POFinish_0400_perMonth
    {
        get
        {
            string mstr_averageTime_POFinish_0400_perMonth = "";
            for (int i = 0; i < averageTime_POFinish_0400_perMonth.Length; i++)
            {
                mstr_averageTime_POFinish_0400_perMonth = mstr_averageTime_POFinish_0400_perMonth + Convert.ToInt16(averageTime_POFinish_0400_perMonth[i]).ToString() + "!";
            }
            return mstr_averageTime_POFinish_0400_perMonth.Substring(0, mstr_averageTime_POFinish_0400_perMonth.Length - 1);
        }
    }
    public static string str_averageTime_POFinish_0481_perMonth
    {
        get
        {
            string mstr_averageTime_POFinish_0481_perMonth = "";
            for (int i = 0; i < averageTime_POFinish_0481_perMonth.Length; i++)
            {
                mstr_averageTime_POFinish_0481_perMonth = mstr_averageTime_POFinish_0481_perMonth + Convert.ToInt16(averageTime_POFinish_0481_perMonth[i]).ToString() + "!";
            }
            return mstr_averageTime_POFinish_0481_perMonth.Substring(0, mstr_averageTime_POFinish_0481_perMonth.Length - 1);
        }
    }
    public static string str_averageTime_POFinish_perMonth
    {
        get
        {
            string mstr_averageTime_POFinish_perMonth = "";
            for (int i = 0; i < averageTime_POFinish_perMonth.Length; i++)
            {
                mstr_averageTime_POFinish_perMonth = mstr_averageTime_POFinish_perMonth + Convert.ToInt16(averageTime_POFinish_perMonth[i]).ToString() + "!";
            }
            return mstr_averageTime_POFinish_perMonth.Substring(0, mstr_averageTime_POFinish_perMonth.Length - 1);
        }
    }
    #endregion

    #region Send to ECharts.aspx
    
    //ECharts1_1:DCR(Delivery Class Reliablity)
    public static string values_ECharts1_1
    {
        get
        {
            //return str_DCR_inside + ";" + str_DCR_outside;

            //return str_DCR + "!" + Convert.ToInt16(((Convert.ToDouble(count_0481_PO_OnTime_YTD + count_0400_PO_OnTime_YTD) / Convert.ToDouble(count_PO_Finished_YTD)) * 100)).ToString() + ";"
            //    + str_DCR_0400 + "!" + Convert.ToInt16(((Convert.ToDouble(count_0400_PO_OnTime_YTD) / Convert.ToDouble(count_0400_PO_Finished_YTD)) * 100)).ToString() + ";"
            //    + str_DCR_0481 + "!" + Convert.ToInt16(((Convert.ToDouble(count_0481_PO_OnTime_YTD) / Convert.ToDouble(count_0481_PO_Finished_YTD)) * 100)).ToString();

            return str_DCR + "!" + ((int)((Convert.ToDouble(count_0481_PO_OnTime_YTD + count_0400_PO_OnTime_YTD) / Convert.ToDouble(count_PO_Finished_YTD)) * 100 + 0.5)).ToString() + ";"
                + str_DCR_0400 + "!" + ((int)((Convert.ToDouble(count_0400_PO_OnTime_YTD) / Convert.ToDouble(count_0400_PO_Finished_YTD)) * 100 + 0.5)).ToString() + ";"
                + str_DCR_0481 + "!" + ((int)((Convert.ToDouble(count_0481_PO_OnTime_YTD) / Convert.ToDouble(count_0481_PO_Finished_YTD)) * 100 + 0.5)).ToString();
        }
    }
    //ECharts1_2:LT(Average Lead Time：CDS)
    public static string values_ECharts1_2
    {
        get
        {
            return str_LT_perMonth + "!" + Convert.ToInt16(LT_YTD).ToString()
                + ";" + str_LT_0400_perMonth + "!" + Convert.ToInt16(LT_0400_YTD).ToString()
                + ";" + str_LT_0481_perMonth + "!" + Convert.ToInt16(LT_0481_YTD).ToString();
            //return str_LT_perMonth + "!" + str_LT_perWeek + "!" + LT_YTD.ToString("F2")
            //    + ";" + str_LT_0400_perMonth + "!" + str_LT_0400_perWeek + "!" + LT_0400_YTD.ToString("F2")
            //    + ";" + str_LT_0481_perMonth + "!" + str_LT_0481_perWeek + "!" + LT_0481_YTD.ToString("F2");
        }
    }
    //ECharts1_3:Reminders for 3Weeks&2Weeks
    public static string values_ECharts1_3
    {
        get
        {
            //return count_Reminder3Weeks_LT + "!" + count_Reminder2Weeks_LT ;
            return count_Reminder3Weeks_LT + "!" + count_Reminder2Weeks_LT + ";"
               + count_FailedMonitor_DC + "!" + count_FailedMonitor_Req;
        }
    }
    //ECharts1_4:Ongoing Pro. Order Monitoring
    public static string values_ECharts1_4
    {
        get
        {
            return count_Ongoing_40 + "!" + count_Ongoing_40_50 + "!" + count_Ongoing_50;
        }
    }
    //ECharts2_1:SO Create perMonth
    public static string values_ECharts2_1
    {
        get
        {
            return str_SO_soCreated_0400_perMonth + "!" + count_SO_soCreated_0400_YTD.ToString() + ";"
                + str_SO_soCreated_0481_perMonth + "!" + count_SO_soCreated_0481_YTD.ToString() + ";"
                + str_NetValue_SO_soCreated_perMonth + "!" + NetValue_SO_soCreated_YTD.ToString("F2");
        }
    }
    //ECharts2_2:Finish POs
    public static string values_ECharts2_2
    {
        get
        {
            return str_PO_Finished_0400_perMonth + "!" + count_PO_Finished_0400_YTD.ToString() + ";"
                + str_PO_Finished_0481_perMonth + "!" + count_PO_Finished_0481_YTD.ToString() + ";"
                + str_PO_ForecastFinished_perMonth + "!" + count_PO_ForecastFinished_YTD.ToString() + ";"
                + str_NetValue_PO_Finished_perMonth + "!" + NetValue_PO_Finished_YTD.ToString("F2") + ";"
                + str_PO_Finished_perMonth + "!" + count_PO_Finished_YTD.ToString();
        }
    }   
    //ECharts2_3:Repertory Monitor
    public static string values_ECharts2_3
    {
        get
        {
            return count_Stored_5 + "!" + count_Stored_5_10 + "!" + count_Stored_10_20 + "!" + count_Stored_20 + "!" + count_Stored + ";"
                + count_Stored_5_0400 + "!" + count_Stored_5_10_0400 + "!" + count_Stored_10_20_0400 + "!" + count_Stored_20_0400 + "!" + count_Stored_0400 + ";"
                + count_Stored_5_0481 + "!" + count_Stored_5_10_0481 + "!" + count_Stored_10_20_0481 + "!" + count_Stored_20_0481 + "!" + count_Stored_0481;
        }
    }
    //ECharts2_4:Ready for Shipment Monitoring
    public static string values_ECharts2_4
    {
        get
        {
            return count_Ready4Shipment_10 + "!" + count_Ready4Shipment_10_20 + "!" + count_Ready4Shipment_20_30 + "!" + count_Ready4Shipment_30 + ";"
                + count_Ready4Shipment_10_0400 + "!" + count_Ready4Shipment_10_20_0400 + "!" + count_Ready4Shipment_20_30_0400 + "!" + count_Ready4Shipment_30_0400 + ";"
                + count_Ready4Shipment_10_0481 + "!" + count_Ready4Shipment_10_20_0481 + "!" + count_Ready4Shipment_20_30_0481 + "!" + count_Ready4Shipment_30_0481;
        }
    }
    //ECharts3_1:PO Create Monitor
    public static string values_ECharts3_1
    {
        get
        {
            return str_PO_soCreated_perMonth + ";" + str_CreatedDelay_perMonth + ";" + (count_PO_soCreated_0400_YTD - count_CreatedDelay_0400_YTD).ToString() + "!" + count_CreatedDelay_0400_YTD.ToString();
        }
    }
    //ECharts3_2:PO Release Monitor
    public static string values_ECharts3_2
    {
        get
        {
            return str_PO_poCreated_perMonth + ";" + str_ReleaseDelay_perMonth + ";" + (count_PO_poCreated_0400_YTD - count_ReleaseDelay_0400_YTD).ToString() + "!" + count_ReleaseDelay_0400_YTD.ToString();
        }
    }
    //ECharts3_3:PO Release->Actual Finish(23WD)
    public static string values_ECharts3_3
    {
        get
        {
            return str_averageTime_POFinish_perMonth + "!" + Convert.ToInt16(averageTime_POFinish_YTD).ToString()
                + ";" + str_averageTime_POFinish_0400_perMonth + "!" + Convert.ToInt16(averageTime_POFinish_0400_YTD).ToString()
                + ";" + str_averageTime_POFinish_0481_perMonth + "!" + Convert.ToInt16(averageTime_POFinish_0481_YTD).ToString();
            
        }
    }
    //ECharts3_4:Actual Finish->ex-plant(2WD)
    public static string values_ECharts3_4
    {
        get
        {
            return str_averageTime_Ex_plant_perMonth + "!" + Convert.ToInt16(averageTime_Ex_plant_YTD).ToString()
                + ";" + str_averageTime_Ex_plant_0400_perMonth + "!" + Convert.ToInt16(averageTime_Ex_plant_0400_YTD).ToString()
                + ";" + str_averageTime_Ex_plant_0481_perMonth + "!" + Convert.ToInt16(averageTime_Ex_plant_0481_YTD).ToString();
        }
    }
    #endregion

    #region Panel1
    //ECharts1-1:DCR(Delivery Class Reliablity)及时交付率【0400+0481】
    public static int count_0400_PO_DLV;
    public static int count_0400_PO_DLV_OnTime;
    public static int count_0400_PO_DLV_Delay;
    public static int count_0400_PO_noDLV;
    public static int count_0400_PO_noDLV_Delay;
    public static int count_0400_PO_noDLV_Warning;
    public static int count_0400_PO_noDLV_OK;
    public static int count_0400_PO_LastFailed;//昨天刚刚超期的订单数
    public static int[] count_0400_PO_OnTime_perMonth = new int[12];
    public static int count_0400_PO_OnTime_YTD;
    public static int[] count_0400_PO_Finished_perMonth = new int[12];
    public static int count_0400_PO_Finished_YTD;

    public static int count_0481_PO_DLV;
    public static int count_0481_PO_DLV_OnTime;
    public static int count_0481_PO_DLV_Delay;
    public static int count_0481_PO_noDLV;
    public static int count_0481_PO_noDLV_Delay;
    public static int count_0481_PO_noDLV_Warning;
    public static int count_0481_PO_noDLV_OK;
    public static int count_0481_PO_LastFailed;//昨天刚刚超期的订单数
    public static int[] count_0481_PO_OnTime_perMonth = new int[12];
    public static int count_0481_PO_OnTime_YTD;
    public static int[] count_0481_PO_Finished_perMonth = new int[12];
    public static int count_0481_PO_Finished_YTD;
    //ECharts1-2:LT(Average Lead Time：CDS)平均交货时间【0400】
    public static double[] LT_perMonth = new double[12];
    public static double[] LT_perWeek = new double[4];
    public static double LT_YTD;

    public static double[] LT_0400_perMonth = new double[12];
    public static double[] LT_0400_perWeek = new double[4];
    public static double LT_0400_YTD;

    public static double[] LT_0481_perMonth = new double[12];
    public static double[] LT_0481_perWeek = new double[4];
    public static double LT_0481_YTD;
    //ECharts1-3:DCR ALERT【0400】Reminders for 3Weeks&2Weeks & Last Failed Order
    public static int count_Reminder3Weeks_LT;
    public static int count_Reminder2Weeks_LT;
    public static int count_Reminder3Weeks_ConfirmedDt;
    public static int count_Reminder2Weeks_ConfirmedDt;
    public static int count_FailedMonitor_DC;//昨天为止失败的订单
    public static int count_FailedMonitor_Req;
    //ECharts1-4:Ongoing Pro. Order【0400】
    public static int count_Ongoing_40;
    public static int count_Ongoing_40_50;
    public static int count_Ongoing_50;
    #endregion

    #region Panel2
    //ECharts2-1:Monthly SO Input【0400+0481】
    public static int[] count_SO_soCreated_perMonth = new int[12];//需要查重，从List<string>中处理一下
    public static int count_SO_soCreated_YTD;//需要查重，从List<string>中处理一下
    public static double[] NetValue_SO_soCreated_perMonth = new double[12];
    public static double NetValue_SO_soCreated_YTD;

    public static int[] count_SO_soCreated_0400_perMonth = new int[12];//需要查重，从List<string>中处理一下
    public static int count_SO_soCreated_0400_YTD;//需要查重，从List<string>中处理一下
    public static double[] NetValue_SO_soCreated_0400_perMonth = new double[12];
    public static double NetValue_SO_soCreated_0400_YTD;

    public static int[] count_SO_soCreated_0481_perMonth = new int[12];//需要查重，从List<string>中处理一下
    public static int count_SO_soCreated_0481_YTD;//需要查重，从List<string>中处理一下
    public static double[] NetValue_SO_soCreated_0481_perMonth = new double[12];
    public static double NetValue_SO_soCreated_0481_YTD;
    //ECharts2-2:Monthly Finished Pro. Order【0400+0481】
    public static int[] count_PO_Finished_perMonth = new int[12];
    public static int count_PO_Finished_YTD;
    public static double[] NetValue_PO_Finished_perMonth = new double[12];
    public static double NetValue_PO_Finished_YTD;
    public static int[] count_PO_ForecastFinished_perMonth = new int[12];
    public static int count_PO_ForecastFinished_YTD;

    public static int[] count_PO_Finished_0400_perMonth = new int[12];
    public static int count_PO_Finished_0400_YTD;
    public static double[] NetValue_PO_Finished_0400_perMonth = new double[12];
    public static double NetValue_PO_Finished_0400_YTD;
    public static int[] count_PO_ForecastFinished_0400_perMonth = new int[12];
    public static int count_PO_ForecastFinished_0400_YTD;

    public static int[] count_PO_Finished_0481_perMonth = new int[12];
    public static int count_PO_Finished_0481_YTD;
    public static double[] NetValue_PO_Finished_0481_perMonth = new double[12];
    public static double NetValue_PO_Finished_0481_YTD;
    public static int[] count_PO_ForecastFinished_0481_perMonth = new int[12];
    public static int count_PO_ForecastFinished_0481_YTD;
    //ECharts2-3:CS Shipment Statistic【YTD 0400+0481】
    public static int count_Stored_5;
    public static int count_Stored_5_10;
    public static int count_Stored_10_20;
    public static int count_Stored_20;
    public static int count_Stored;//已完成无ShipmentStartOn 

    public static int count_Stored_5_0400;
    public static int count_Stored_5_10_0400;
    public static int count_Stored_10_20_0400;
    public static int count_Stored_20_0400;
    public static int count_Stored_0400;//已完成无ShipmentStartOn 

    public static int count_Stored_5_0481;
    public static int count_Stored_5_10_0481;
    public static int count_Stored_10_20_0481;
    public static int count_Stored_20_0481;
    public static int count_Stored_0481;//已完成无ShipmentStartOn 
    //ECharts2-4:Ready fot Shipment Monitoring【YTD 0400+0481】
    public static int count_Ready4Shipment_10;
    public static int count_Ready4Shipment_10_20;
    public static int count_Ready4Shipment_20_30;
    public static int count_Ready4Shipment_30;

    public static int count_Ready4Shipment_10_0400;
    public static int count_Ready4Shipment_10_20_0400;
    public static int count_Ready4Shipment_20_30_0400;
    public static int count_Ready4Shipment_30_0400;

    public static int count_Ready4Shipment_10_0481;
    public static int count_Ready4Shipment_10_20_0481;
    public static int count_Ready4Shipment_20_30_0481;
    public static int count_Ready4Shipment_30_0481;
    #endregion

    #region Panel3 【0400】E2E Process Monitoring
    //ECharts3-1:SO->PO Creation Monitoring(1WD)
    public static int[] count_PO_soCreated_perMonth = new int[12];
    public static int[] count_CreatedDelay_perMonth = new int[12];
    public static int count_PO_soCreated_YTD;
    public static int count_CreatedDelay_YTD;

    public static int[] count_PO_soCreated_0400_perMonth = new int[12];
    public static int[] count_CreatedDelay_0400_perMonth = new int[12];
    public static int count_PO_soCreated_0400_YTD;
    public static int count_CreatedDelay_0400_YTD;
    //ECharts3-2:PO Creation->PO Release Monitoring(5WD)
    public static int[] count_PO_poCreated_perMonth = new int[12];
    public static int[] count_ReleaseDelay_perMonth = new int[12];
    public static int count_PO_poCreated_YTD;
    public static int count_ReleaseDelay_YTD;

    public static int[] count_PO_poCreated_0400_perMonth = new int[12];
    public static int[] count_ReleaseDelay_0400_perMonth = new int[12];
    public static int count_PO_poCreated_0400_YTD;
    public static int count_ReleaseDelay_0400_YTD;
    //ECharts3-3:PO Release->Actual finish Monitoring(23WD)【Average】
    public static double[] averageTime_POFinish_perMonth = new double[12];
    public static double averageTime_POFinish_YTD;
    public static double[] averageTime_POFinish_0400_perMonth = new double[12];
    public static double averageTime_POFinish_0400_YTD;
    public static double[] averageTime_POFinish_0481_perMonth = new double[12];
    public static double averageTime_POFinish_0481_YTD;
    //ECharts3-4:Actual finish->Ex-plant Monitoring(2WD)【Average】
    public static double[] averageTime_Ex_plant_perMonth = new double[12];
    public static double averageTime_Ex_plant_YTD;
    public static double[] averageTime_Ex_plant_0400_perMonth = new double[12];
    public static double averageTime_Ex_plant_0400_YTD;
    public static double[] averageTime_Ex_plant_0481_perMonth = new double[12];
    public static double averageTime_Ex_plant_0481_YTD;
    #endregion

    public static void CleanCounter()
    {
        #region Panel1
        //ECharts1-1:DCR(Delivery Class Reliablity)及时交付率【0400+0481】
        count_0400_PO_DLV = 0;
        count_0400_PO_DLV_OnTime = 0;
        count_0400_PO_DLV_Delay = 0;
        count_0400_PO_noDLV = 0;
        count_0400_PO_noDLV_Delay = 0;
        count_0400_PO_noDLV_Warning = 0;
        count_0400_PO_noDLV_OK = 0;
        count_0400_PO_LastFailed = 0;//昨天刚刚超期的订单数
        count_0400_PO_OnTime_perMonth = new int[12];
        count_0400_PO_OnTime_YTD = 0;
        count_0400_PO_Finished_perMonth = new int[12];
        count_0400_PO_Finished_YTD = 0;

        count_0481_PO_DLV = 0;
        count_0481_PO_DLV_OnTime = 0;
        count_0481_PO_DLV_Delay = 0;
        count_0481_PO_noDLV = 0;
        count_0481_PO_noDLV_Delay = 0;
        count_0481_PO_noDLV_Warning = 0;
        count_0481_PO_noDLV_OK = 0;
        count_0481_PO_LastFailed = 0;//昨天刚刚超期的订单数
        count_0481_PO_OnTime_perMonth = new int[12];
        count_0481_PO_OnTime_YTD = 0;
        count_0481_PO_Finished_perMonth = new int[12];
        count_0481_PO_Finished_YTD = 0;
        //ECharts1-2:LT(Average Lead Time：CDS)平均交货时间【0400】
        LT_perMonth = new double[12];
        LT_perWeek = new double[4];
        LT_YTD = 0;

        LT_0400_perMonth = new double[12];
        LT_0400_perWeek = new double[4];
        LT_0400_YTD = 0;

        LT_0481_perMonth = new double[12];
        LT_0481_perWeek = new double[4];
        LT_0481_YTD = 0;
        //ECharts1-3:DCR ALERT【0400】Reminders for 3Weeks&2Weeks & Last Failed Order
        count_Reminder3Weeks_LT = 0;
        count_Reminder2Weeks_LT = 0;
        count_Reminder3Weeks_ConfirmedDt = 0;
        count_Reminder2Weeks_ConfirmedDt = 0;
        count_FailedMonitor_DC = 0;//昨天为止失败的订单
        count_FailedMonitor_Req = 0;
        //ECharts1-4:Ongoing Pro. Order【0400】
        count_Ongoing_40 = 0;
        count_Ongoing_40_50 = 0;
        count_Ongoing_50 = 0;
        #endregion

        #region Panel2
        //ECharts2-1:Monthly SO Input【0400+0481】
        count_SO_soCreated_perMonth = new int[12];//需要查重，从List<string>中处理一下
        count_SO_soCreated_YTD = 0;//需要查重，从List<string>中处理一下
        NetValue_SO_soCreated_perMonth = new double[12];
        NetValue_SO_soCreated_YTD = 0;

        count_SO_soCreated_0400_perMonth = new int[12];//需要查重，从List<string>中处理一下
        count_SO_soCreated_0400_YTD = 0;//需要查重，从List<string>中处理一下
        NetValue_SO_soCreated_0400_perMonth = new double[12];
        NetValue_SO_soCreated_0400_YTD = 0;

        count_SO_soCreated_0481_perMonth = new int[12];//需要查重，从List<string>中处理一下
        count_SO_soCreated_0481_YTD = 0;//需要查重，从List<string>中处理一下
        NetValue_SO_soCreated_0481_perMonth = new double[12];
        NetValue_SO_soCreated_0481_YTD = 0;
        //ECharts2-2:Monthly Finished Pro. Order【0400+0481】
        count_PO_Finished_perMonth = new int[12];
        count_PO_Finished_YTD = 0;
        NetValue_PO_Finished_perMonth = new double[12];
        NetValue_PO_Finished_YTD = 0;
        count_PO_ForecastFinished_perMonth = new int[12];
        count_PO_ForecastFinished_YTD = 0;

        count_PO_Finished_0400_perMonth = new int[12];
        count_PO_Finished_0400_YTD = 0;
        NetValue_PO_Finished_0400_perMonth = new double[12];
        NetValue_PO_Finished_0400_YTD = 0;
        count_PO_ForecastFinished_0400_perMonth = new int[12];
        count_PO_ForecastFinished_0400_YTD = 0;

        count_PO_Finished_0481_perMonth = new int[12];
        count_PO_Finished_0481_YTD = 0;
        NetValue_PO_Finished_0481_perMonth = new double[12];
        NetValue_PO_Finished_0481_YTD = 0;
        count_PO_ForecastFinished_0481_perMonth = new int[12];
        count_PO_ForecastFinished_0481_YTD = 0;
        //ECharts2-3:CS Shipment Statistic【YTD 0400+0481】
        count_Stored_5 = 0;
        count_Stored_5_10 = 0;
        count_Stored_10_20 = 0;
        count_Stored_20 = 0;
        count_Stored = 0;//已完成无ShipmentStartOn 
        count_Stored_5_0400 = 0;
        count_Stored_5_10_0400 = 0;
        count_Stored_10_20_0400 = 0;
        count_Stored_20_0400 = 0;
        count_Stored_0400 = 0;//已完成无ShipmentStartOn 
        count_Stored_5_0481 = 0;
        count_Stored_5_10_0481 = 0;
        count_Stored_10_20_0481 = 0;
        count_Stored_20_0481 = 0;
        count_Stored_0481 = 0;//已完成无ShipmentStartOn 
        //ECharts2-4:Ready fot Shipment Monitoring【YTD 0400+0481】
        count_Ready4Shipment_10 = 0;
        count_Ready4Shipment_10_20 = 0;
        count_Ready4Shipment_20_30 = 0;
        count_Ready4Shipment_30 = 0;
        count_Ready4Shipment_10_0400 = 0;
        count_Ready4Shipment_10_20_0400 = 0;
        count_Ready4Shipment_20_30_0400 = 0;
        count_Ready4Shipment_30_0400 = 0;
        count_Ready4Shipment_10_0481 = 0;
        count_Ready4Shipment_10_20_0481 = 0;
        count_Ready4Shipment_20_30_0481 = 0;
        count_Ready4Shipment_30_0481 = 0;
        #endregion

        #region Panel3 【0400】E2E Process Monitoring
        //ECharts3-1:SO->PO Creation Monitoring(1WD)
        count_PO_soCreated_perMonth = new int[12];
        count_CreatedDelay_perMonth = new int[12];
        count_PO_soCreated_YTD = 0;
        count_CreatedDelay_YTD = 0;

        count_PO_soCreated_0400_perMonth = new int[12];
        count_CreatedDelay_0400_perMonth = new int[12];
        count_PO_soCreated_0400_YTD = 0;
        count_CreatedDelay_0400_YTD = 0;
        //ECharts3-2:PO Creation->PO Release Monitoring(5WD)
        count_PO_poCreated_perMonth = new int[12];
        count_ReleaseDelay_perMonth = new int[12];
        count_PO_poCreated_YTD = 0;
        count_ReleaseDelay_YTD = 0;

        count_PO_poCreated_0400_perMonth = new int[12];
        count_ReleaseDelay_0400_perMonth = new int[12];
        count_PO_poCreated_0400_YTD = 0;
        count_ReleaseDelay_0400_YTD = 0;
        //ECharts3-3:PO Release->Actual finish Monitoring(23WD)【Average】
        averageTime_POFinish_perMonth = new double[12];
        averageTime_POFinish_YTD = 0;
        averageTime_POFinish_0400_perMonth = new double[12];
        averageTime_POFinish_0400_YTD = 0;
        averageTime_POFinish_0481_perMonth = new double[12];
        averageTime_POFinish_0481_YTD = 0;
        //ECharts3-4:Actual finish->Ex-plant Monitoring(2WD)【Average】
        averageTime_Ex_plant_perMonth = new double[12];
        averageTime_Ex_plant_YTD = 0;
        averageTime_Ex_plant_0400_perMonth = new double[12];
        averageTime_Ex_plant_0400_YTD = 0;
        averageTime_Ex_plant_0481_perMonth = new double[12];
        averageTime_Ex_plant_0481_YTD = 0;
        #endregion


    }
}