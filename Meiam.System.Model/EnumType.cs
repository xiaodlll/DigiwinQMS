/*
* ==============================================================================
*
* FileName: EnumType.cs
* Created: 2020/6/4 10:40:13
* Author: Meiam
* Description: 
*
* ==============================================================================
*/
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace Meiam.System.Model
{
    /// <summary>
    /// 枚举扩展属性
    /// </summary>
    public static class EnumExtension
    {
        /// <summary>
        /// 获得枚举提示文本
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string GetEnumText(this Enum obj)
        {
            Type type = obj.GetType();
            FieldInfo field = type.GetField(obj.ToString());
            TextAttribute attribute = (TextAttribute)field.GetCustomAttribute(typeof(TextAttribute));
            return attribute.Value;
        }
    }

    public class TextAttribute : Attribute
    {
        public TextAttribute(string value)
        {
            Value = value;
        }

        public string Value { get; set; }
    }

    public enum StatusCodeType
    {
        [Text("请求(或处理)成功")]
        Success = 200,

        [Text("内部请求出错")]
        Error = 500,

        [Text("访问请求未授权! 当前 SESSION 失效, 请重新登陆")]
        Unauthorized = 401,

        [Text("请求参数不完整或不正确")]
        ParameterError = 400,

        [Text("您无权进行此操作，请求执行已拒绝")]
        Forbidden = 403,

        [Text("找不到与请求匹配的 HTTP 资源")]
        NotFound = 404,

        [Text("HTTP请求类型不合法")]
        HttpMehtodError = 405,

        [Text("HTTP请求不合法,请求参数可能被篡改")]
        HttpRequestError = 406,

        [Text("该URL已经失效")]
        URLExpireError = 407,
    }


    public enum SourceType
    {
        /// <summary>
        /// 后台程序
        /// </summary>
        [Text("Web")]
        Web,

        /// <summary>
        /// 微信小程序
        /// </summary>
        [Text("MiniProgram")]
        MiniProgram,

        /// <summary>
        /// 福利小程序
        /// </summary>
        [Text("Perk")]
        Perk,
    }

    /// <summary>
    /// 项目阶段
    /// </summary>
    public enum ProjectStage
    {
        New = 0,//新建
        Pending = 1,//待报价
        Quotation = 2,//报价中
        CostAcc = 3,//成本核算
        Audit = 4,//价格审核
    }

    /// <summary>
    /// 项目状态
    /// </summary>
    public enum ProjectStatus
    {
        Running = 0,//进行中
        Suspend = 1,//暂停
        Close = 2,//终止
        Filed = 3,//归档
        Delete = 4,//删除
    }
}
