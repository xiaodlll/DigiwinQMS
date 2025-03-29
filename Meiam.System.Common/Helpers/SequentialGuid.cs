using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Schema;
using System.Xml.Serialization;
using System.Xml;

namespace Meiam.System.Common.Helpers
{
    public class SequentialGuid
    {
        /// <summary>
        /// 生成一个有序的 GUID，并保证字符串形式按顺序排列。
        /// </summary>
        /// <returns>有序 GUID 的字符串形式</returns>
        public static string Generate()
        {
            // 获取当前时间戳（以毫秒为单位）
            DateTime now = DateTime.UtcNow;
            long timestamp = now.Ticks / 10000; // 将 Ticks 转换为毫秒（1 Tick = 100 纳秒）

            // 创建一个随机 GUID 的后 10 字节
            byte[] randomBytes = Guid.NewGuid().ToByteArray();
            byte[] orderedGuidBytes = new byte[16];

            // 将时间戳的前 6 字节写入 GUID 的前 6 字节
            byte[] timestampBytes = BitConverter.GetBytes(timestamp);
            if (BitConverter.IsLittleEndian)
            {
                Array.Reverse(timestampBytes); // 确保是大端字节序
            }
            Array.Copy(timestampBytes, timestampBytes.Length - 6, orderedGuidBytes, 0, 6);

            // 填充剩余的 10 字节为随机字节
            Array.Copy(randomBytes, 6, orderedGuidBytes, 6, 10);

            // 创建 GUID 对象
            Guid orderedGuid = new Guid(orderedGuidBytes);

            // 返回字符串形式的 GUID
            return orderedGuid.ToString().ToUpper();
        }
    }
}
