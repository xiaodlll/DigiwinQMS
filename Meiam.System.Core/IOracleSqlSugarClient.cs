using SqlSugar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Meiam.System.Core {


    public interface IOracleSqlSugarClient : ISqlSugarClient {
        // 保持空接口，仅用于依赖注入区分
    }

    public class OracleSqlSugarClient : SqlSugarScope, IOracleSqlSugarClient {
        public OracleSqlSugarClient(ConnectionConfig config, Action<SqlSugarClient> configAction)
            : base(config, configAction) {
        }
    }
}
