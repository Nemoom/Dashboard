using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
///ProcessMonitor 的摘要说明
/// </summary>
public class ProcessMonitor
{
	public ProcessMonitor()
	{
		//
		//TODO: 在此处添加构造函数逻辑
		//
	}

    public static string mPO;
    public static string mSO;
    //SKA/FKA
    public static string mSKA_FKA;
    //KA Name
    public static string mKA_Name;
    //订单交付工厂代号
    public static DLV_Plant mDLV_Plant;
    //该订单的净值
    public static double mNetValue;
    //Excel中的实际时间节点
    public static ActualTime mActualTime;
    //Excel中的预计时间节点
    public static EstimatedTime mEstimatedTime;

    //订单完成状态
    public static PO_Status mPO_Status;
    //依据报价承诺给客户多少个工作日交付计算出的具体交付日期
    public static DateTime Date_QuotationLT { get { return HolidayHelper.GetInstance().GetReckonDate(mActualTime.Date_SO_CreatedOn, mEstimatedTime.QuotationLT, true); } }
    //SO创建到PO创建的时间间隔(WD)
    public static int Gap_SO2PO { get { return Math.Abs(HolidayHelper.GetInstance().GetWorkDayNum(mActualTime.Date_SO_CreatedOn, mActualTime.Date_PO_CreatedOn, false)); } }
    //PO创建到PO释放的时间间隔(WD)
    public static int Gap_PO2Release { get { return Math.Abs(HolidayHelper.GetInstance().GetWorkDayNum(mActualTime.Date_PO_CreatedOn, mActualTime.Date_ActualReleaseDate, false)); } }
    //PO释放到完成生产入库的时间间隔(WD)
    public static int Gap_Release2Finish { get { return Math.Abs(HolidayHelper.GetInstance().GetWorkDayNum(mActualTime.Date_ActualReleaseDate, mActualTime.Date_ActualFinishDate, false)); } }
    //完成生产入库到起运发货的时间间隔(WD)
    public static int Gap_Finish2Shipment { get { return HolidayHelper.GetInstance().GetWorkDayNum(mActualTime.Date_ActualFinishDate, mActualTime.Date_ShipmentStartOn, false); } }
    
    //SO已创建到今天的时间间隔(WD)
    public static int Gap_SO2Today { get { return Math.Abs(HolidayHelper.GetInstance().GetWorkDayNum(mActualTime.Date_SO_CreatedOn, DateTime.Today, false)); } }
    //PO已创建到今天的时间间隔(WD)
    public static int Gap_PO2Today { get { return Math.Abs(HolidayHelper.GetInstance().GetWorkDayNum(mActualTime.Date_PO_CreatedOn, DateTime.Today, false)); } }
    //PO已释放到今天的时间间隔(WD)
    public static int Gap_Release2Today { get { return Math.Abs(HolidayHelper.GetInstance().GetWorkDayNum(mActualTime.Date_ActualReleaseDate, DateTime.Today, false)); } }
    //已完成生产入库但未发货订单，完成生产入库至今的时间间隔(WD)
    public static int Gap_Finish2Today { get { return Math.Abs(HolidayHelper.GetInstance().GetWorkDayNum(mActualTime.Date_ActualFinishDate, DateTime.Today, false)); } }
    //已完成订单Lead Time(CDS)
    public static int LT
    {
        get
        {
            if (mPO_Status == PO_Status.DLV || mPO_Status == PO_Status.PO_Finished)
            {
                return (mActualTime.Date_ActualFinishDate - mActualTime.Date_SO_CreatedOn).Days + 1;
            }
            else
            {
                return 0;
            }
        }
    }

    //未完成订单Lead Time(CDS)
    public static int LT_Ongoing
    {
        get
        {
            if (mPO_Status == PO_Status.DLV || mPO_Status == PO_Status.PO_Finished)
            {
                return 0;
            }
            else
            {
                return (DateTime.Today - mActualTime.Date_SO_CreatedOn).Days;
            }
        }
    }  

    public struct ActualTime 
    {
        public DateTime Date_SO_CreatedOn;
        public DateTime Date_PO_CreatedOn;
        public DateTime Date_ActualReleaseDate;
        public DateTime Date_ActualFinishDate;
        public DateTime Date_1stGRforPO;
        public DateTime Date_ShipmentStartOn;
    }
    public struct EstimatedTime 
    {
        public DateTime Date_1stConfirmedDt;
        public int QuotationLT;
        public DateTime Date_RequestDate;
        public DateTime Date_BasicEndDate;
        public DateTime Date_2ndConfirmedDt;
    }

    public static void CleanDate()
    {
        mActualTime.Date_SO_CreatedOn = new DateTime();
        mActualTime.Date_PO_CreatedOn = new DateTime();
        mActualTime.Date_ActualReleaseDate = new DateTime();
        mActualTime.Date_ActualFinishDate = new DateTime();
        mActualTime.Date_1stGRforPO = new DateTime();
        mActualTime.Date_ShipmentStartOn = new DateTime();

        mEstimatedTime.Date_1stConfirmedDt = new DateTime();
        mEstimatedTime.Date_RequestDate = new DateTime();
        mEstimatedTime.Date_BasicEndDate = new DateTime();
        mEstimatedTime.Date_2ndConfirmedDt = new DateTime();
        mEstimatedTime.QuotationLT = 0;
    }
}

public enum PO_Status 
{
    SO_Created,
    PO_Created,
    PO_Released,
    PO_Finished,
    DLV
}
public enum DLV_Plant
{
    P_0400,
    P_0481    
}