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
    
    public partial class CSSPItemPath
    {
        public int CSSPItemPathID { get; set; }
        public int CSSPItemID { get; set; }
        public Nullable<int> CSSPItemParentID { get; set; }
        public int Level { get; set; }
        public string Path { get; set; }
        public bool HasChild { get; set; }
        public System.DateTime LastModifiedDate { get; set; }
        public int ModifiedByID { get; set; }
        public bool IsActive { get; set; }
    
        public virtual CSSPItem CSSPItem { get; set; }
        public virtual CSSPItem CSSPItem1 { get; set; }
    }
}