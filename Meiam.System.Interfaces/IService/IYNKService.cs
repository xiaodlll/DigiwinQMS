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
        Task<ApiResponse> ProcessLotNoticeAsync(List<LotNoticeRequest> request);

        //收料检验结果回传ERP
        List<LotNoticeResultRequest> GetQmsLotNoticeResultRequest();
        void CallBackQmsLotNoticeResult(LotNoticeResultRequest request);
        Task<ApiResponse> GetAOIInspectInfoByDocCodeAsync(INSPECT_REQCODE input);
        Task<ApiResponse> GetAOIProgressDataByDocCodeAsync(INSPECT_REQCODE input);
        Task<ApiResponse> ProcessUploadAOIDataAsync(List<InspectAoi> input);
        Task<ApiResponse> ProcessUploadAOIImageDataAsync(List<InspectImageAoi> input);
        Task<ApiResponse> ProcessYNKInpectProcessDataAsync(List<INSPECT_PROGRESSDto> input);
    }
}
