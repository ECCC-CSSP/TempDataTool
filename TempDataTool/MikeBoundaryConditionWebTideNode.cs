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
    
    public partial class MikeBoundaryConditionWebTideNode
    {
        public int MikeBoundaryConditionWebTideNodeID { get; set; }
        public int MikeBoundaryConditionID { get; set; }
        public int Ordinal { get; set; }
        public double Lat { get; set; }
        public double Long { get; set; }
    
        public virtual MikeBoundaryCondition MikeBoundaryCondition { get; set; }
    }
}
