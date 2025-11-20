using Meiam.System.Model;
using Meiam.System.Model.Dto;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Meiam.System.Interfaces.IService
{
    public interface IYNKService : IBaseService<INSPECT_TENSILE_D>
    {
        /// <summary>
        /// ERP收料通知单
        /// </summary>
        /// <param name="request">收料通知单请求数据</param>
        /// <returns>处理结果</returns>
        Task<ApiResponse> ProcessLotNoticeAsync(List<LotNoticeRequestYNK> request);

        //金蝶云登录接口,获取KDSVCSessionId 的值
        Task<ERPLoginResponseYNK> LoginAsync();
        Task<ERPLoginResponseYNK> LoginAsync(string username, string password, string acctID, int lcid = 2052);

        //收料检验结果回传ERP
        List<LotNoticeResultRequestYNK> GetQmsLotNoticeResultRequest();
        void CallBackQmsLotNoticeResult(LotNoticeResultRequestYNK request);

        Task<ApiResponse> GetAOIInspectInfoByDocCodeAsync(INSPECT_REQCODE input);
        Task<ApiResponse> GetAOIProgressDataByDocCodeAsync(INSPECT_REQCODE input);
        Task<ApiResponse> ProcessUploadAOIDataAsync(List<InspectAoi> input);
        Task<ApiResponse> ProcessUploadAOIImageDataAsync(List<InspectImageAoi> input);
        Task<ApiResponse> ProcessYNKInpectProcessDataAsync(List<INSPECT_PROGRESSDto> input);
        List<AttachmentResultRequestYNK> GetAttachmentResultRequest();
        void CallBackAttachmentResult(AttachmentResultRequestYNK item);
        Task<ApiResponse> GetInspectionRecordReportDataAsync(INSPECT_REQCODE input);
    }
}
