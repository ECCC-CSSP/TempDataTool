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
    
    public partial class WQMSubsecDefPrecipStation
    {
        public WQMSubsecDefPrecipStation()
        {
            this.WQMDefPrecipStations = new HashSet<WQMDefPrecipStation>();
        }
    
        public int SubsecDefPrecipStationID { get; set; }
        public int SubsectorID { get; set; }
        public Nullable<System.DateTime> StartDate { get; set; }
        public Nullable<System.DateTime> EndDate { get; set; }
        public Nullable<System.DateTime> LastModifiedDate { get; set; }
        public Nullable<int> ModifiedByID { get; set; }
        public Nullable<bool> IsActive { get; set; }
    
        public virtual ICollection<WQMDefPrecipStation> WQMDefPrecipStations { get; set; }
        public virtual WQMSubsector WQMSubsector { get; set; }
    }
}