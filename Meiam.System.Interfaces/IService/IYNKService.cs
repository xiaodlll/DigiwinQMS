using Meiam.System.Model;
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
    }
}
