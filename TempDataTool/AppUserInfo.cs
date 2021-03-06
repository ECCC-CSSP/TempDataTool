//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace TempDataTool
{
    using System;
    using System.Collections.Generic;
    
    public partial class AppUserInfo
    {
        public AppUserInfo()
        {
            this.AppTasks = new HashSet<AppTask>();
            this.CSSPItemsUserAuthorizations = new HashSet<CSSPItemsUserAuthorization>();
            this.CSSPItemTypesUserAuthorizations = new HashSet<CSSPItemTypesUserAuthorization>();
        }
    
        public int AppUserInfoID { get; set; }
        public System.Guid UserID { get; set; }
        public string LoginEmail { get; set; }
        public string FullName { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Initial { get; set; }
        public string WebName { get; set; }
        public string Address { get; set; }
        public string TelWork { get; set; }
        public string TelHome { get; set; }
        public string TelCell { get; set; }
        public string WorkEmail { get; set; }
        public string HomeEmail { get; set; }
        public Nullable<System.DateTime> DateOfBirth { get; set; }
        public Nullable<System.DateTime> LastModifiedDate { get; set; }
        public Nullable<int> ModifiedByID { get; set; }
        public Nullable<bool> IsActive { get; set; }
    
        public virtual ICollection<AppTask> AppTasks { get; set; }
        public virtual aspnet_Users aspnet_Users { get; set; }
        public virtual ICollection<CSSPItemsUserAuthorization> CSSPItemsUserAuthorizations { get; set; }
        public virtual ICollection<CSSPItemTypesUserAuthorization> CSSPItemTypesUserAuthorizations { get; set; }
    }
}
