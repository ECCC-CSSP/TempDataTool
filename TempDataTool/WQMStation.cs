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
    
    public partial class WQMStation
    {
        public WQMStation()
        {
            this.WQMSamples = new HashSet<WQMSample>();
        }
    
        public int StationID { get; set; }
        public int SubsectorID { get; set; }
        public string StationNumber { get; set; }
        public string StationName { get; set; }
        public string StationType { get; set; }
        public string SortOrder { get; set; }
        public int Zone { get; set; }
        public double Easting { get; set; }
        public double Northing { get; set; }
        public decimal Lat { get; set; }
        public decimal Long { get; set; }
        public bool Active { get; set; }
        public System.DateTime Updated { get; set; }
        public Nullable<System.DateTime> LastModifiedDate { get; set; }
        public Nullable<int> ModifiedByID { get; set; }
        public Nullable<bool> IsActive { get; set; }
    
        public virtual ICollection<WQMSample> WQMSamples { get; set; }
        public virtual WQMSubsector WQMSubsector { get; set; }
    }
}