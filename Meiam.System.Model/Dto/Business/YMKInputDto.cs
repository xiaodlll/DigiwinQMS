using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Meiam.System.Model.Dto
{
    public class InspectAoi
    {
        public string DOC_CODE { get; set; }
        public string INSPECT_PROGRESSID { get; set; }
        public string INSPECT_AOIID { get; set; }
        public DateTime INSPECT_AOICREATEDATE { get; set; }
        public string INSPECT_AOICREATEUSER { get; set; }
        public string TENID { get; set; }
        public string MainSN { get; set; }
        public string PanelSN { get; set; }
        public string PanelID { get; set; }
        public string ModelName { get; set; }
        public string Side { get; set; }
        public string MachineName { get; set; }
        public string CustomerName { get; set; }
        public string Operator { get; set; }
        public string Programer { get; set; }
        public string InspectionDate { get; set; }
        public string BeginTime { get; set; }
        public string EndTime { get; set; }
        public string CycleTimeSec { get; set; }
        public string InspectionBatch { get; set; }
        public string ReportResult { get; set; }
        public string ConfirmedResult { get; set; }
        public string TotalComponent { get; set; }
        public string ReportFailComponent { get; set; }
        public string ComfirmedFailComponent { get; set; }
        public string ComponentName { get; set; }
        public string LibraryModel { get; set; }
        public string PN { get; set; }
        public string Package { get; set; }
        public string Angle { get; set; }
        public string NGReportResult { get; set; }
        public string ReportResultCode { get; set; }
        public string NGConfirmedResult { get; set; }
        public string ConfirmedResultCode { get; set; }
    }

    public class InspectImageAoi
    {
        public string DOC_CODE { get; set; }
        public string ImageName { get; set; }
        public string ImageData { get; set; }
    }

    public class INSPECT_PERSONNELDATA
    {
        public string SumType { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string MeterialNames { get; set; }
    }
}
