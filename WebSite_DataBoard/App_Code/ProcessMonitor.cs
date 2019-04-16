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

    public static DateTime Date_QuotationLT { get { return HolidayHelper.GetInstance().GetReckonDate(mActualTime.Date_SO_CreatedOn, mEstimatedTime.QuotationLT, false); } }
    public static int Gap_SO2PO { get { return (mActualTime.Date_PO_CreatedOn - mActualTime.Date_SO_CreatedOn).Days; } }
    public static int Gap_PO2Release { get { return (mActualTime.Date_ActualReleaseDate - mActualTime.Date_PO_CreatedOn).Days; } }
    public static int Gap_Finish2Shipment{ get { return (mActualTime.Date_ShipmentStartOn - mActualTime.Date_ActualFinishDate).Days; } }
    public static ActualTime mActualTime;
    public static EstimatedTime mEstimatedTime;

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

}