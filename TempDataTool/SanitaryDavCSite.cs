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
    
    public partial class SanitaryDavCSite
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public SanitaryDavCSite()
        {
            this.SanitaryDavCObs = new HashSet<SanitaryDavCOb>();
        }
    
        public int SanitaryDavCSiteID { get; set; }
        public Nullable<int> siteid { get; set; }
        public string Locator { get; set; }
        public Nullable<int> Site { get; set; }
        public Nullable<int> Zone { get; set; }
        public Nullable<double> easting { get; set; }
        public Nullable<double> northing { get; set; }
        public string Datum { get; set; }
        public Nullable<double> latitude { get; set; }
        public Nullable<double> longitude { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SanitaryDavCOb> SanitaryDavCObs { get; set; }
    }
}
