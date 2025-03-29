/*
* ==============================================================================
*
* FileName: UserSession.cs
* Created: 2020/5/19 14:39:24
* Author: Meiam
* Description: 
*
* ==============================================================================
*/
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Meiam.System.Model
{
    public class UserSessionVM
    {

        /// <summary>
        /// 描述 : 用户账号 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "用户账号")]
        public string UserID { get; set; }

        /// <summary>
        /// 描述 : 用户名称 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "用户名称")]
        public string UserName { get; set; }

        /// <summary>
        /// 描述 : 用户昵称 
        /// 空值 : True
        /// 默认 : 
        /// </summary>
        [Display(Name = "用户昵称")]
        public string NickName { get; set; }

        /// <summary>
        /// 描述 : 邮箱 
        /// 空值 : True
        /// 默认 : 
        /// </summary>
        [Display(Name = "邮箱")]
        public string Email { get; set; }

        /// <summary>
        /// 描述 : 性别 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "性别")]
        public string Sex { get; set; }

        /// <summary>
        /// 描述 : 照片 
        /// 空值 : True
        /// 默认 : 
        /// </summary>
        [Display(Name = "头像地址")]
        public string AvatarUrl { get; set; }

        /// <summary>
        /// 描述 : 工号 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "工号")]
        public string EmpCode { get; set; }

        /// <summary>
        /// 描述 : 部门 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "部门")]
        public string DepartmentName { get; set; }

        /// <summary>
        /// 描述 : 手机号码 
        /// 空值 : True
        /// 默认 : 
        /// </summary>
        [Display(Name = "手机号码")]
        public string Phone { get; set; }

        /// <summary>
        /// 描述 : 备注 
        /// 空值 : True
        /// 默认 : 
        /// </summary>
        [Display(Name = "备注")]
        public string Remark { get; set; }

        /// <summary>
        /// 描述 : 创建时间 
        /// 空值 : True
        /// 默认 : 
        /// </summary>
        [Display(Name = "创建时间")]
        public DateTime? CreateTime { get; set; }

        /// <summary>
        /// 描述 : 是否启用 
        /// 空值 : False
        /// 默认 : 1
        /// </summary>
        [Display(Name = "是否启用")]
        public bool Enabled { get; set; }

        /// <summary>
        /// 描述 : 单用户模式 
        /// 空值 : False
        /// 默认 : 0
        /// </summary>
        [Display(Name = "单用户模式")]
        public bool OneSession { get; set; }

        /// <summary>
        /// 描述 : 来源   
        /// 空值 : False    
        /// 默认 :   
        /// </summary>
        [Display(Name = "来源")]
        public string Source { get; set; }

        /// <summary>
        /// 描述 : 持续时间   
        /// 空值 : False    
        /// 默认 :   
        /// </summary>
        [Display(Name = "持续时间")]
        public int KeepHours { get; set; }

        /// <summary>
        /// 描述 : 超级管理员   
        /// 空值 : False    
        /// 默认 :   
        /// </summary>
        [Display(Name = "超级管理员")]
        public bool Administrator { get; set; } = false;

        /// <summary>
        /// 描述 : 系统权限   
        /// 空值 : False    
        /// 默认 :   
        /// </summary>
        [Display(Name = "系统权限")]
        public List<string> UserPower { get; set; }

        /// <summary>
        /// 描述 : 系统角色   
        /// 空值 : False    
        /// 默认 :   
        /// </summary>
        [Display(Name = "系统角色")]
        public List<string> UserRole { get; set; }
        
    }
}
