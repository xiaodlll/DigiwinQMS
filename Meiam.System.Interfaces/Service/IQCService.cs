//------------------------------------------------------------------------------
// <auto-generated>
//     此代码已从模板生成手动更改此文件可能导致应用程序出现意外的行为。
//     如果重新生成代码，将覆盖对此文件的手动更改。
//     author MEIAM
// </auto-generated>
//------------------------------------------------------------------------------
using Meiam.System.Model;
using Meiam.System.Model.Dto;
using System.Collections.Generic;
using System.Threading.Tasks;
using SqlSugar;
using System.Linq;
using System;
using System.IO;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot;
using OxyPlot.Core.Drawing;
using OxyPlot.Legends;
using Microsoft.Data.SqlClient;
using System.Data;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;
using Microsoft.Extensions.Options;
using DotNetCore.CAP;
using Mapster.Utils;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using SqlSugar.Extensions;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System.Reflection.Emit;

namespace Meiam.System.Interfaces
{
    public class IQCService : BaseService<INSPECT_TENSILE>, IIQCService {

        public IQCService(IUnitOfWork unitOfWork) : base(unitOfWork) {
        }

        #region CustomInterface 
        public void SaveToInspectDetail(IEnumerable<INSPECT_TENSILE_D> inspectData){
            try {
                // 开启事务
                Db.Ado.BeginTran();
                // 批量插入数据
                Db.Insertable<INSPECT_TENSILE_D>(inspectData).ExecuteCommand();

                List<INSPECT_TENSILE_D_R> listInspectDataR = new List<INSPECT_TENSILE_D_R>();
                foreach (var detail in inspectData) {
                    // 去除字符串首尾的花括号
                    string trimmedInput = detail.Y_AXIS.Trim('{', '}');
                    // 按逗号分割字符串
                    string[] numberStrings = trimmedInput.Split(',');
                    // 将分割后的字符串数组转换为双精度浮点数数组
                    decimal[] numbers = Array.ConvertAll(numberStrings, decimal.Parse);
                    decimal max = numbers.Max();
                    decimal min = numbers.Min();
                    decimal avg = numbers.Average();

                    var drItem = new INSPECT_TENSILE_D_R() {
                        INSPECT_TENSILE_D_RID = Guid.NewGuid().ToString(),
                        INSPECT_TENSILE_DID = detail.INSPECT_TENSILE_DID,
                        MaxValue = max,
                        MinValue = min,
                        AvgValue = avg
                    };
                    listInspectDataR.Add(drItem);
                }

                Db.Insertable<INSPECT_TENSILE_D_R>(listInspectDataR).ExecuteCommand();

                // 提交事务
                Db.Ado.CommitTran();
            }
            catch (Exception) {
                // 回滚事务
                Db.Ado.RollbackTran();
                throw;
            }
        }
        public byte[] GetInspectImage(List<INSPECT_TENSILE_D> listInspectData){
            byte[] result = null;

            // 创建一个 PlotModel 对象，它代表整个图表
            var plotModel = new PlotModel {
                Title = "拉力机检测图"
            };

            // 创建一个线性坐标轴作为 X 轴
            plotModel.Axes.Add(new LinearAxis {
                Position = AxisPosition.Bottom,
                Title = "变形(mm)",
                Minimum = 0,
                Maximum = 1.15
            });

            // 创建一个线性坐标轴作为 Y 轴
            plotModel.Axes.Add(new LinearAxis {
                Position = AxisPosition.Left,
                Title = "力(gt)",
                Minimum = 0,
                Maximum = 200
            });

            int i = 1;
            Random random = new Random();
            foreach (var detail in listInspectData) {
                string trimmedInput = detail.X_AXIS.Trim('{', '}');
                string[] numberStrings = trimmedInput.Split(',');
                double[] dataArraysX = Array.ConvertAll(numberStrings, double.Parse);

                trimmedInput = detail.Y_AXIS.Trim('{', '}');
                numberStrings = trimmedInput.Split(',');
                double[] dataArraysY = Array.ConvertAll(numberStrings, double.Parse);

                // 生成随机颜色
                byte r = (byte)random.Next(256);
                byte g = (byte)random.Next(256);
                byte b = (byte)random.Next(256);
                OxyColor randomColor = OxyColor.FromRgb(r, g, b);

                var lineSeries = new LineSeries {
                    Title = $"Plot{i++}",
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 1,
                    MarkerStroke = randomColor,
                MarkerFill = OxyColors.White,
                Color = randomColor
                };

                for (int j = 0; j < dataArraysX.Length; j++) {
                    lineSeries.Points.Add(new DataPoint(dataArraysX[j], dataArraysY[j]));
                }

                // 将曲线添加到图表模型中
                plotModel.Series.Add(lineSeries);
            }

            // 保存图表为 PNG 图片
            var exporter = new PngExporter { Width = 800, Height = 600, Resolution = 100 };
            using (MemoryStream stream = new MemoryStream()) {
                exporter.Export(plotModel, stream);
                stream.Position = 0;
                result = stream.ToArray();
                stream.Flush();
            }
            return result;
        }

        public bool ExistScanDoc(string docType, string peopleId){
            string sql = @"SELECT COUNT(*) FROM SCANDOC WHERE DOCTYPE=@DOCTYPE AND PEOPLEID=@PEOPLEID";

            // 定义参数
            var parameters = new SugarParameter[]
            {
                new SugarParameter("@DOCTYPE", docType),
                new SugarParameter("@PEOPLEID", peopleId)
            };

            // 执行 SQL 命令
            int count = Db.Ado.GetInt(sql, parameters);
            return count > 0;
        }

        public void SaveToScanDoc(string docType, byte[] fileContents, string scandocName, string peopleId) {
            FileInfo file = new FileInfo(scandocName);
            if (!file.Directory.Exists) {
                file.Directory.Create();
            }
            File.WriteAllBytes(scandocName, fileContents);

            string sql = @"
            INSERT INTO SCANDOC (TENID, SCANDOCID, SCANDOCCODE, SCANDOCNAME, DOCTYPE, PEOPLEID, createdate, SCANDOC_user)
            VALUES ('001', @SCANDOCID, @SCANDOCID, @SCANDOCNAME, @DOCTYPE, @PEOPLEID, CONVERT(VARCHAR(20), GETDATE(), 120), @PEOPLEID)";

            // 定义参数
            var parameters = new SugarParameter[]
            {
                new SugarParameter("@SCANDOCID", Guid.NewGuid().ToString()),
                new SugarParameter("@SCANDOCNAME", scandocName),
                new SugarParameter("@DOCTYPE", docType),
                new SugarParameter("@PEOPLEID", peopleId)
            };

            // 执行 SQL 命令
            Db.Ado.ExecuteCommand(sql, parameters);
        }
        #endregion

        #region GetCPKfile
        public byte[] GetCPKfile(string INSPECT_DEV2ID,string userName) 
        {
            
            string INSPECT_CODE;//检验单号
            string INSPECT_PUR; //检验来源

            //测试
            //GET_INSPECT_LIST("INSPECT_ZONE_021", "IQC_2025030003", "IQC");

            # region 一．获得检验单号和检验来源
            string sql = @"SELECT Top 1 ISNULL(INSPECT_DEV2.INSPECT_CODE,'') AS INSPECT_CODE,ISNULL(INSPECT_DEV2.INSPECT_PUR,'') As INSPECT_PUR    
                        FROM INSPECT_DEV2 
                        LEFT JOIN INSPECT_FLOW ON INSPECT_FLOW.INSPECT_FLOWID=INSPECT_DEV2.INSPECT_FLOWID
                        LEFT JOIN COLUM002 ON COLUM002.COLUM002ID=INSPECT_DEV2.COLUM002ID
                        WHERE INSPECT_DEV2ID=@INSPECT_DEV2ID";
            // 定义参数
            var parameters = new SugarParameter[]
            {
                new SugarParameter("@INSPECT_DEV2ID", INSPECT_DEV2ID)
            };

            // 执行 SQL 命令
            var dataTable = Db.Ado.GetDataTable(sql, parameters);

            if (dataTable.Rows.Count > 0)
            {
                INSPECT_CODE = dataTable.Rows[0]["INSPECT_CODE"].ToString();
                INSPECT_PUR = dataTable.Rows[0]["INSPECT_PUR"].ToString();
            }
            else
            {
                throw new Exception("未获取到检验单号和检验来源");
            }
            #endregion

            #region 二．判断检验单是否已经完成
            //如果 检验单号 不为空 并且 检验来源 不为空 @动态表名 =“INSPECT_”+@INSPECT_PUR
            //@动态表面ID栏位 = @动态表名 +“ID”
            //    SELECT STATE FROM @动态表名 WHERE @动态表面ID栏位 = @INSPECT_CODE
            //    如果 STATE =“已完成”
            //    返回错误：”检验单已完成，无法再次产生检验报告” 提出API
            string table = string.Empty;
            string surfaceId = string.Empty;

            if (!string.IsNullOrEmpty(INSPECT_CODE) && !string.IsNullOrEmpty(INSPECT_PUR))
            {
                table = "INSPECT_" + INSPECT_PUR;
                surfaceId = table + "ID";
                string state = "";

                sql = @$"SELECT STATE FROM {table} WHERE {surfaceId}=@INSPECT_CODE";
                // 定义参数
                parameters = new SugarParameter[]
                {
                    new SugarParameter("@INSPECT_CODE", INSPECT_CODE)
                };
                // 执行 SQL 命令
                dataTable = Db.Ado.GetDataTable(sql, parameters);

                if (dataTable.Rows.Count > 0)
                {
                    state = dataTable.Rows[0]["STATE"].ToString();
                }
                else
                {
                    throw new Exception("未获取到检验单状态");
                }

                if (state == "已完成")
                {
                    throw new Exception("检验单已完成，无法再次产生检验报告");
                }
            }
            else
            {
                throw new Exception("检验单号或检验来源为空");
            }
            #endregion

            #region 三．执行存储过程  
            //执行存储过程
            //如果 返回值前两位 =“错误”，则 退出API，将返回值返回
            // 执行 SQL 命令
            string INPECT_CODE = Db.Ado.GetString(@$"EXEC DEV2_GET_INPECT_CODE '{INSPECT_DEV2ID}','COC_ATTR_001'");
            if (INPECT_CODE.Contains("错误"))
            {
                throw new Exception($"DEV2_GET_INPECT_CODE获取异常：{INPECT_CODE}");
            }

            #endregion

            #region 四．重新获得主档资料
            string COLUM002ID = "";
            string ITEMID = "";
            string LOTID = "";
            string INSPECT_FLOWID = "";

            sql = @"SELECT Top 1 ISNULL(INSPECT_DEV2.INSPECT_CODE, '') AS INSPECT_CODE
                    ,ISNULL(INSPECT_DEV2.INSPECT_PUR, '') AS INSPECT_PUR  -- 检验来源 IQC OQC
                    ,ISNULL(INSPECT_FLOW.ITEMID, '') AS ITEMID-- 物料编码(有个低代码待办)
                    ,ISNULL(INSPECT_DEV2.LOTID, '') AS LOTID-- 批次号
                    ,ISNULL(INSPECT_DEV2.DOC_CODE, '') AS DOC_CODE-- 来源单号
                    ,ISNULL(INSPECT_DEV2.COLUM002ID, '') AS COLUM002ID
                    ,ISNULL(COLUM002.COC_ATTR, '') AS COC_ATTR--特殊设定
                    ,ISNULL(INSPECT_DEV2.INSPECT_FLOWID, '') AS INSPECT_FLOWID
                    FROM INSPECT_DEV2
                    LEFT JOIN INSPECT_FLOW ON INSPECT_FLOW.INSPECT_FLOWID = INSPECT_DEV2.INSPECT_FLOWID
                    LEFT JOIN COLUM002 ON COLUM002.COLUM002ID = INSPECT_DEV2.COLUM002ID
                    WHERE INSPECT_DEV2ID = @INSPECT_DEV2ID";
            parameters = new SugarParameter[]
            {
                new SugarParameter("@INSPECT_DEV2ID", INSPECT_DEV2ID)
            };

            // 执行 SQL 命令
            dataTable = Db.Ado.GetDataTable(sql, parameters);

            if (dataTable.Rows.Count > 0)
            {
                COLUM002ID = dataTable.Rows[0]["COLUM002ID"].ToString();
                ITEMID = dataTable.Rows[0]["ITEMID"].ToString();
                LOTID = dataTable.Rows[0]["LOTID"].ToString();
                INSPECT_FLOWID = dataTable.Rows[0]["INSPECT_FLOWID"].ToString();
            }

            #endregion

            #region 五．获得一些资料 为第七、第八步使用
            //得到 @COLUM002结果集
            sql = @"SELECT * FROM COLUM002 where COLUM002ID = @COLUM002ID";
            parameters = new SugarParameter[]
            {
                new SugarParameter("@COLUM002ID", COLUM002ID)
            };
            // 执行 SQL 命令
            int count_COLUM002 = Db.Ado.GetInt(sql, parameters);

            //得到 @第一个样本ID
            sql = @"SELECT TOP 1 SAMPLEID FROM INSPECT_2D  WHERE INSPECT_DEV2ID = @INSPECT_DEV2ID";
            parameters = new SugarParameter[]
            {
                new SugarParameter("@INSPECT_DEV2ID", INSPECT_DEV2ID)
            };
            int SAMPLEID = Db.Ado.GetInt(sql, parameters);

            //如果 INSPECT_2D异常结果集 记录数> 0 则
            //返回错误：“当前设备原始LOCATION资料和当前选择的检验项目的检验内容不一致”
            //得到 INSPECT_2D异常结果集
            sql = @"SELECT * FROM  INSPECT_2D
                LEFT JOIN COLUM001 ON INSPECT_2D.LOCATION = COLUM001.COLUM001NAME
                WHERE INSPECT_DEV2ID = @INSPECT_DEV2ID  AND SAMPLEID = @SAMPLEID
                AND COLUM001.COLUM002ID = @COLUM002ID  AND COLUM001.COLUM001CODE IS NOT NULL
                ORDER BY LOCATION";
            parameters = new SugarParameter[]
            {
                new SugarParameter("@INSPECT_DEV2ID", INSPECT_DEV2ID),
                new SugarParameter("@SAMPLEID", SAMPLEID),
                new SugarParameter("@COLUM002ID", COLUM002ID)
            };
            // 执行 SQL 命令
            int count_INSPECT_2D = Db.Ado.GetInt(sql, parameters);

            #endregion

            #region 六 .获取应检样本数 @检验批次数量LOT_QTY
            //六．获得应检样本数
            //由@COLUM002ID获得应检样本数
            //执行存储过程：GET_INPECT_CNT @COLUM002ID, @检验批次数量LOT_QTY
            //@应检样本数 = 结果集CNT列的值

            int lot_Qyt = 0;
            sql = @$"SELECT top 1 ISNULL(LOT_QTY,0)  FROM {table} WHERE {surfaceId}=@INSPECT_CODE";
            // 定义参数
            parameters = new SugarParameter[]
            {
                new SugarParameter("@INSPECT_CODE", INSPECT_CODE)
            };
            // 执行 SQL 命令
            dataTable = Db.Ado.GetDataTable(sql, parameters);
            if (dataTable.Rows.Count > 0)
            {
                lot_Qyt = int.Parse(dataTable.Rows[0][0].ToString());
            }
            else
            {
                throw new Exception("未获取到检验批次数量LOT_QTY");
            }

            //获得应检样本量
            int inspect_Qyt = Db.Ado.GetInt(@$"EXEC GET_INPECT_CNT  '{COLUM002ID}','{lot_Qyt}'");

            #endregion

            #region 七．进一步检验
            if (count_INSPECT_2D > 0)
            {
                throw new Exception("当前设备原始LOCATION资料和当前选择的检验项目的检验内容不一致");
            }
            #endregion

            #region 八．更新检验内容的CODE给 INSPECT_2D
            //如果 @COLUM002结果集 记录数 > 0 并且 @INSPECT_2D异常结果集 记录数 > 0
            if (count_COLUM002>0 && count_INSPECT_2D>0)
            {
                //（通过LOCATION关联检验内容，更新设备原始记录）
                sql = @$"UPDATE INSPECT_2D SET
                        COLUM001CODE = COLUM001.COLUM001CODE,ADD_VALUE = COLUM001.ADD_VALUE
                        FROM INSPECT_2D
                        LEFT JOIN COLUM001 ON INSPECT_2D.LOCATION = COLUM001.COLUM001NAME
                        WHERE INSPECT_DEV2ID = @INSPECT_DEV2ID
                        AND SAMPLEID = @SAMPLEID
                        AND COLUM001.COLUM002ID = @COLUM002ID
                        AND COLUM001.COLUM001CODE IS NOT NULL";
                // 定义参数
                parameters = new SugarParameter[]
                {
                    new SugarParameter("@INSPECT_DEV2ID", INSPECT_DEV2ID),
                    new SugarParameter("@SAMPLEID", SAMPLEID),
                    new SugarParameter("@COLUM002ID", COLUM002ID)
                    
                };

                // 执行 SQL 命令
                Db.Ado.ExecuteCommand(sql, parameters);
            }
            #endregion

            #region 九．将 INSPECT_2D 的检验内容（检验位置）传递给QMS
            //如果 @COLUM002结果集 记录数 = 0 或者为NULL
            if (count_COLUM002 == 0)
            {
                #region 1.根据@COLUM002ID 产生检验内容（COLUM001）
                string CREATEUSER = "system";
                string TENID = "001";
                string COLUM001ID = "";
                string @CUSTOMID = "";
                string RE = "";
                string MAX_VALUE = "";  //--INSPECT_2D.VALUE2
                string STD_VALUE = "";  //标准值: INSPECT_2D.VALUE1
                string COLUM001NAME = "";//--INSPECT_2D.LOCATION
                string INSPECT_LEVELID = "";
                string OPTIONS = "";
                string COLUM001CODE = "";//--第一个A01，第二个A02，依次类推
                string MIN_VALUE = "";  //INSPECT_2D.VALUE3
                string REMARK1 = "二次元";
                string AC = "";
                string INSPECT_AQLCODE = "";
                string REMARK = "";
                string REMARK2 = "";
                string COLUM0A10 = "Num";
                string INSPECT_PLANID = "c6cae8ea-24e0-4fbe-ac6e-775843549e5b";
                string INSPECT_2DID;
                //根据 @COLUM002ID 产生检验内容（COLUM001）
                //--得到第一个样本的所有检验内容
                sql = @$"SELECT VALUE2,VALUE1,LOCATION,VALUE3,INSPECT_2DID,* FROM  INSPECT_2D
                WHERE  INSPECT_DEV2ID = @INSPECT_DEV2ID  AND SAMPLEID = @SAMPLEID";

                parameters = new SugarParameter[]
                {
                    new SugarParameter("@INSPECT_DEV2ID", INSPECT_DEV2ID),
                    new SugarParameter("@SAMPLEID", SAMPLEID),
                };
                // 执行 SQL 命令
                dataTable = Db.Ado.GetDataTable(sql, parameters);

                if (dataTable.Rows.Count > 0)
                {
                    for (int i=0; i < dataTable.Rows.Count; i++)
                    {
                        STD_VALUE = dataTable.Rows[i]["VALUE1"].ToString(); 
                        MAX_VALUE = dataTable.Rows[i]["VALUE2"].ToString(); 
                        MIN_VALUE = dataTable.Rows[i]["VALUE3"].ToString(); 
                        COLUM001NAME = dataTable.Rows[i]["LOCATION"].ToString(); 
                        COLUM001CODE = "A" + (i + 1).ToString("D2");
                        INSPECT_2DID= dataTable.Rows[i]["INSPECT_2DID"].ToString();
                        COLUM001ID = Guid.NewGuid().ToString();

                        sql = @$"INSERT INTO COLUM001(COLUM001CREATEDATE, COLUM001CREATEUSER, TENID, COLUM001ID, 
                        CUSTOMID, RE, MAX_VALUE, STD_VALUE, COLUM001NAME, INSPECT_LEVELID, OPTIONS, 
                        COLUM001CODE, MIN_VALUE, COLUM002ID, REMARK1, AC, INSPECT_AQLCODE, REMARK, REMARK2, 
                        COLUM0A10, INSPECT_PLANID) 
                        VALUES( 
                        CONVERT(VARCHAR(20), GETDATE(), 120), @CREATEUSER, @TENID, @COLUM001ID, @CUSTOMID, 
                        @RE, @MAX_VALUE, @STD_VALUE, @COLUM001NAME, @INSPECT_LEVELID, @OPTIONS, @COLUM001CODE, 
                        @MIN_VALUE, @COLUM002ID, @REMARK1, @AC, @INSPECT_AQLCODE, @REMARK, @REMARK2, @COLUM0A10, @INSPECT_PLANID)";

                        parameters = new SugarParameter[]
                        {
                            new SugarParameter("@CREATEUSER", CREATEUSER),
                            new SugarParameter("@TENID", TENID),
                            new SugarParameter("@COLUM001ID", COLUM001ID),
                            new SugarParameter("@CUSTOMID", CUSTOMID),
                            new SugarParameter("@RE", RE),
                            new SugarParameter("@MAX_VALUE", MAX_VALUE),
                            new SugarParameter("@STD_VALUE", STD_VALUE),
                            new SugarParameter("@COLUM001NAME", COLUM001NAME),
                            new SugarParameter("@INSPECT_LEVELID", INSPECT_LEVELID),
                            new SugarParameter("@OPTIONS", OPTIONS),
                            new SugarParameter("@COLUM001CODE", COLUM001CODE),
                            new SugarParameter("@MIN_VALUE", MIN_VALUE),
                            new SugarParameter("@COLUM002ID", COLUM002ID),
                            new SugarParameter("@REMARK1", REMARK1),
                            new SugarParameter("@AC", AC),
                            new SugarParameter("@INSPECT_AQLCODE", INSPECT_AQLCODE),
                            new SugarParameter("@REMARK", REMARK),
                            new SugarParameter("@REMARK2", REMARK2),
                            new SugarParameter("@COLUM0A10", COLUM0A10),
                            new SugarParameter("@INSPECT_PLANID", INSPECT_PLANID),
                        };
                        // 执行 SQL 命令
                        Db.Ado.ExecuteCommand(sql, parameters);

                        #region 2.回写A01的编码给原始资料
                        sql = @$"UPDATE INSPECT_2D SET COLUM001CODE = @COLUM001CODE WHERE INSPECT_2DID = @INSPECT_2DID";
                        parameters = new SugarParameter[]
                        {
                            new SugarParameter("@COLUM001CODE", COLUM001CODE),
                            new SugarParameter("@INSPECT_2DID", INSPECT_2DID),
                        };
                        // 执行 SQL 命令
                        Db.Ado.ExecuteCommand(sql, parameters);

                        //Db.Ado.ExecuteCommand(@$"UPDATE INSPECT_2D SET COLUM001CODE ='{COLUM001CODE}' WHERE INSPECT_2DID = '{INSPECT_2DID}'");
                        #endregion
                    }
                }
                #endregion

                
            }
            #endregion

            #region 十.将INSPECT_2D实际值传入QMS
            //1.获得INSPECT_2D实际测量的样本数量
            int actCNT = Db.Ado.GetInt(@$"SELECT COUNT(DISTINCT SAMPLEID) AS CNT FROM INSPECT_2D WHERE INSPECT_DEV2ID ='{INSPECT_DEV2ID}'");
            //2.如果 @CNT > 0 则
            if (actCNT > 0)
            {
                //2.1 删除QMS中的记录：@COLUM002ID @INSPECT_CODE
                sql = @$" DELETE INSPECT_ZONE WHERE COLUM002ID = @COLUM002ID AND INSPECTCODE = @INSPECT_CODE";
                parameters = new SugarParameter[]
                {
                    new SugarParameter("@COLUM002ID", COLUM002ID),
                    new SugarParameter("@INSPECT_CODE", INSPECT_CODE),
                };
                Db.Ado.ExecuteCommand(sql, parameters);
                //2.2 开始同步记录
                //循环每个SAMPLEID
                sql = @$" SELECT  SAMPLEID  FROM INSPECT_2D WHERE  INSPECT_DEV2ID=@INSPECT_DEV2ID GROUP BY SAMPLEID";
                parameters = new SugarParameter[]
                {
                    new SugarParameter("@INSPECT_DEV2ID", INSPECT_DEV2ID),
                };
                dataTable = Db.Ado.GetDataTable(sql, parameters);

                if (dataTable.Rows.Count > 0)
                {
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        SAMPLEID = Convert.ToInt32(dataTable.Rows[i]["SAMPLEID"]);

                        #region 1 循环结果集
                        sql = @$"SELECT COLUM001CODE  AS 检验内容编码, MAX(VALUE)  AS 检验值
                            FROM INSPECT_2D WHERE INSPECT_DEV2ID = @INSPECT_DEV2ID
                            AND SAMPLEID = @SAMPLEID
                            GROUP BY COLUM001CODE
                            ORDER BY COLUM001CODE";
                        parameters = new SugarParameter[]
                        {
                            new SugarParameter("@INSPECT_DEV2ID", INSPECT_DEV2ID),
                            new SugarParameter("@SAMPLEID", SAMPLEID),
                        };
                        var dataTable1 = Db.Ado.GetDataTable(sql, parameters);

                        string sel_Col = "";
                        string sel_VALUES = "";

                        if (dataTable1.Rows.Count > 0)
                        {
                            for (int j = 0; j < dataTable1.Rows.Count; j++)
                            {
                                sel_Col += "," + dataTable1.Rows[j]["检验内容编码"].ToString();
                                sel_VALUES += "," +"'"+ dataTable1.Rows[j]["检验值"].ToString() +"'";
                            }
                        }
                        #endregion
                        #region 2.插入检验结果
                        sql = @$"INSERT INSPECT_ZONE(INSPECT_ZONECREATEUSER, INSPECT_ZONECREATEDATE, INSPECT_ZONEID, INSPECTTYPE, COLUM002ID, 
                            CUSTOM_ITEMID, LOTNO, INSPECTCODE, PCSCODE, ISAUTO{sel_Col}) 
                            VALUES(
                            '{userName}',CONVERT(varchar(20), GETDATE(), 120), '{Guid.NewGuid().ToString()}','{INSPECT_PUR}', '{COLUM002ID}',
                            '{ITEMID}','{LOTID}','{INSPECT_CODE}','{i.ToString()}','1'{sel_VALUES})";

                        Db.Ado.ExecuteCommand(sql);

                        #endregion
                    }

                }
            }

            #endregion

            #region 十一．让QMS产生随机值
            //1.如果 @应检样本数 > @CNT 则 @产生样本数量 = @应检样本数 - @CNT
            int generateQty = 0;

            if (actCNT < inspect_Qyt)
            {
                generateQty = inspect_Qyt - actCNT;
            }
            GET_INSPECT_RANK(COLUM002ID, INSPECT_CODE, inspect_Qyt, INSPECT_PUR, userName);

            #endregion

            #region 十二．产生【CPK-扩展项目】的随机值
            //1.得到CPK - 扩展项结果集
            sql = @$"SELECT COLUM002ID FROM COLUM002 WHERE INSPECT_FLOWID = @INSPECT_FLOWID WHERE COC_ATTR ='COC_ATTR_002'";
            parameters = new SugarParameter[]
            {
                new SugarParameter("@INSPECT_FLOWID", INSPECT_FLOWID)
            };
            // 执行 SQL 命令
            dataTable = Db.Ado.GetDataTable(sql, parameters);
            //2.循环CPK - 扩展项结果集
            if (dataTable.Rows.Count > 0)
            {
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    COLUM002ID = dataTable.Rows[i]["COLUM002ID"].ToString();
                    //1.获得应检样本量
                    inspect_Qyt = Db.Ado.GetInt(@$"EXEC GET_INPECT_CNT  '{COLUM002ID}','{lot_Qyt}'");
                    //2.产生检验随机值
                    //传入参数：
                    //@COLUM002ID(检验项目)--CPK - 扩展项结果集.@COLUM002ID
                    //@INSPECT_CODE(检验单号)(前文获取过）
                    //@产生样本数--@CPK - 扩展项目应检样本数
                    //@INSPECT_PUR--检验类别(前文获取过）
                    //@userName
                    GET_INSPECT_RANK(COLUM002ID, INSPECT_CODE, inspect_Qyt, INSPECT_PUR, userName);
                }
            }
            #endregion

            return null;
        }
        #endregion

        #region GET_INSPECT_RANK
        public void GET_INSPECT_RANK(string COLUM002ID, string INSPECT_CODE, int intSampleCount, string INSPECT_PUR, string userName)
        {
            #region 0.删除QMS中的随机记录
            string sql = @"DELETE FROM INSPECT_ZONE WHERE COLUM002ID = @COLUM002ID AND INSPECTCODE = @INSPECT_CODE AND ISAUTO = '1'";
            var parameters = new SugarParameter[]
            {
                new SugarParameter("@COLUM002ID", COLUM002ID),
                new SugarParameter("@INSPECT_CODE", INSPECT_CODE),
            };
            Db.Ado.ExecuteCommand(sql, parameters);
            #endregion

            #region 2.获得已存在记录数：
            int exist_Qyt = 0;
            sql = @$"SELECT COUNT(1) FROM INSPECT_ZONE WHERE INSPECTCODE = @INSPECTCODE";
            // 定义参数
            parameters = new SugarParameter[]
            {
                new SugarParameter("@INSPECTCODE", INSPECT_CODE)
            };
            // 执行 SQL 命令
            var dataTable = Db.Ado.GetDataTable(sql, parameters);
            if (dataTable.Rows.Count > 0)
            {
                exist_Qyt = int.Parse(dataTable.Rows[0][0].ToString());
            }
            #endregion

            #region 3.获得@ITEMID，@LOTID
            //@动态表名 =“INSPECT_”+@INSPECT_PUR
            //@动态表面ID栏位 = @动态表名 +“ID” 
            //SELECT @ITEMID = ITEMID，@LOTID = LOTID FROM @动态表名 WHERE @动态表面ID栏位 = @INSPECT_CODE
            string table = string.Empty;
            string surfaceId = string.Empty;
            string itemId = string.Empty;
            string lotId = string.Empty;

            table = "INSPECT_" + INSPECT_PUR;
            surfaceId = table + "ID";

            sql = @$"SELECT top 1 ITEMID,LOTID FROM {table} WHERE {surfaceId}=@INSPECT_CODE";
            // 定义参数
            parameters = new SugarParameter[]
            {
                new SugarParameter("@INSPECT_CODE", INSPECT_CODE)
            };
            // 执行 SQL 命令
            dataTable = Db.Ado.GetDataTable(sql, parameters);

            if (dataTable.Rows.Count > 0)
            {
                itemId = dataTable.Rows[0]["ITEMID"].ToString();
                lotId = dataTable.Rows[0]["lotId"].ToString();
            }
            #endregion

            #region 1.获得结果集A：
            sql = @$"SELECT ISNULL(STD_VALUE,0) as STD_VALUE,ISNULL(MIN_VALUE,0) as MIN_VALUE,
		        ISNULL(MAX_VALUE,0) as MAX_VALUE,ISNULL(ADD_VALUE,0) as ADD_VALUE,ISNULL(COLUM001CODE,'') as COLUM001CODE
                FROM COLUM001 WHERE COLUM002ID = @COLUM002ID";
            parameters = new SugarParameter[]
            {
                new SugarParameter("@COLUM002ID", COLUM002ID)
            };
            // 执行 SQL 命令
            dataTable = Db.Ado.GetDataTable(sql, parameters);
            #endregion

            # region 4 循环样本：
            double std_VALUE = 0;
            double min_VALUE =0;
            double max_VALUE = 0;
            double add_VALUE = 0;
            string COLUM001CODE = "";
            double lower_Value = 0; //下限值
            double upper_Value = 0; //上限值
            double act_Value = 0; //生成随机数 实际值
            string sel_Col = "";
            string sel_VALUES = "";
            string sampleId = "";

            int sampleCount = 1;
            while (sampleCount <= intSampleCount)
            {
                #region 1.循环结果集A
                //获得：@标准值 结果集A.STD_VALUE
                //@上公差 结果集A.MIN_VALUE
                //@下公差 结果集A.MAX_VALUE
                //@上下公差余量 = 结果集A.ADD_VALUE
                //得到：@实际值：
                //如：标准值 = 1.08，上公差0.01 下公差 0.03  参数上下公差余量 = 0.001
                //则随机范围是 （1.08 - 0.03 + 0.001~1.08 + 0.01 - 0.001）
                //BEGIN
                //     @变量SELECT +=”,”+结果集A.COLUM001CODE
                //     @变量VALUES +=”,”+实际值
                //END

                if (dataTable.Rows.Count > 0)
                {
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        std_VALUE = Convert.ToDouble(dataTable.Rows[i]["STD_VALUE"]);
                        min_VALUE = Convert.ToDouble(dataTable.Rows[i]["MIN_VALUE"]);
                        max_VALUE = Convert.ToDouble(dataTable.Rows[i]["MAX_VALUE"]);
                        add_VALUE = Convert.ToDouble(dataTable.Rows[i]["ADD_VALUE"]);
                        COLUM001CODE = dataTable.Rows[i]["COLUM001CODE"].ToString();

                        lower_Value = std_VALUE - max_VALUE + add_VALUE;
                        upper_Value = std_VALUE + min_VALUE - add_VALUE;

                        //在范围内生成随机数  
                        Random random = new Random();
                        act_Value = lower_Value + (random.NextDouble() * (upper_Value - lower_Value));

                        sel_Col += "," + COLUM001CODE;
                        sel_VALUES += "," + "'" + act_Value.ToString() + "'";
                    }
                }

                #endregion

                # region 2.插入检验结果
                //样本ID：@已存在记录数 + 循环次数
                sampleId = exist_Qyt+sampleCount.ToString();

                sql = @$"INSERT INSPECT_ZONE(NSPECT_ZONECREATEUSER,INSPECT_ZONECREATEDATE,INSPECT_ZONEID,INSPECTTYPE,COLUM002ID,
                            CUSTOM_ITEMID,LOTNO,INSPECTCODE,PCSCODE,COLUM001ID,ISAUTO{sel_Col}) 
                            VALUES(
                            '{userName}',CONVERT(varchar(20), GETDATE(), 120), '{Guid.NewGuid().ToString()}','{INSPECT_PUR}', '{COLUM002ID}',
                            '{itemId}','{lotId}','{INSPECT_CODE}','{sampleId}','1'{sel_VALUES})";

                Db.Ado.ExecuteCommand(sql);

                #endregion

                sampleCount++;
            }
            #endregion 
        }

        #endregion

        #region GET_INSPECT_LIST

        //@COLUM002ID       --需要查询的检验项目
        //@INSPECT_CODE     --检验单号
        //@INSPECT_PUR      --检验类别
        public DataTable GET_INSPECT_LIST(string COLUM002ID,string INSPECT_CODE, string INSPECT_PUR)
        {
            DataTable dataTableR = new DataTable(); // 创建 DataTable 实例
            # region 1.获得结果集A +2.循环结果集A：

            string sql = @$"SELECT COLUM001CODE,STD_VALUE,MIN_VALUE,MAX_VALUE,REMARK1
                FROM COLUM001 WHERE COLUM002ID = @COLUM002ID and Colum001CODE like'A%'";
            var parameters = new SugarParameter[]
            {
                new SugarParameter("@COLUM002ID", COLUM002ID)
            };
            // 执行 SQL 命令
            var dataTable = Db.Ado.GetDataTable(sql, parameters);

            string COLUM001CODE = string.Empty;
            string sel_Col = string.Empty;
            string sel_ColB = string.Empty;

            if (dataTable.Rows.Count > 0)
            {
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    COLUM001CODE = dataTable.Rows[i]["COLUM001CODE"].ToString();

                    sel_Col += "," + COLUM001CODE;
                    sel_ColB += "," + COLUM001CODE.Replace('A', 'B');
                }
            }

            #endregion

            #region 3.组装SQL
            sql = @$"SELECT INSPECT_ZONECREATEUSER, INSPECT_ZONECREATEDATE, INSPECT_ZONEID, INSPECTTYPE, 
                    COLUM002ID, CUSTOM_ITEMID AS ITEMID, LOTNO, INSPECTCODE, PCSCODE,ISAUTO
                    {sel_Col} {sel_ColB}
                    FROM INSPECT_ZONE
                    WHERE COLUM002ID = @COLUM002ID AND INSPECTCODE = @INSPECT_CODE AND INSPECTTYPE = @INSPECT_PUR 
                    ORDER BY PCSCODE";
            parameters = new SugarParameter[]
            {
                new SugarParameter("@COLUM002ID", COLUM002ID),
                new SugarParameter("@INSPECT_CODE", INSPECT_CODE),
                new SugarParameter("@INSPECT_PUR", INSPECT_PUR),
            };
            // 执行 SQL 命令
            dataTableR = Db.Ado.GetDataTable(sql, parameters);

            #endregion

            return dataTableR; // 返回填充的数据表  
            //返回的所有列名以A开头的列数量，就是区块1的列数
            //返回的所有列名以A开头的列名，就是区块1第15行的列名
            //返回的数据就是32~63行区域的内容（行数依照查询的结果展示）
        }
        #endregion

        #region 
        public DataTable GET_STD_VALUE_LIST(string COLUM002ID, string INSPECT_CODE, string INSPECT_PUR)
        {
            DataTable dataTableR = new DataTable(); // 创建 DataTable 实例

            string sql = @$"SELECT COLUM001CODE,STD_VALUE,MIN_VALUE,MAX_VALUE,REMARK1
                    FROM COLUM001 WHERE COLUM002ID = @COLUM002ID and Colum001CODE like'A%'";
            var parameters = new SugarParameter[]
            {
                new SugarParameter("@COLUM002ID", COLUM002ID)
            };
            // 执行 SQL 命令
            dataTableR = Db.Ado.GetDataTable(sql, parameters);

            return dataTableR; // 返回填充的数据表  
        }
        #endregion
    }

}
