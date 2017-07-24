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
    
    public partial class MikeParameter
    {
        public int MikeParameterID { get; set; }
        public int MikeScenarioID { get; set; }
        public Nullable<double> WindSpeed { get; set; }
        public Nullable<double> WindDirection { get; set; }
        public Nullable<double> DecayFactorPerDay { get; set; }
        public Nullable<bool> DecayIsConstant { get; set; }
        public Nullable<double> DecayFactorAmplitude { get; set; }
        public Nullable<int> ResultFrequencyInMinutes { get; set; }
        public Nullable<double> AmbientTemperature { get; set; }
        public Nullable<double> AmbientSalinity { get; set; }
        public Nullable<double> ManningNumber { get; set; }
        public Nullable<System.DateTime> LastModifiedDate { get; set; }
        public Nullable<int> ModifiedByID { get; set; }
        public Nullable<bool> IsActive { get; set; }
    
        public virtual MikeScenario MikeScenario { get; set; }
    }
}
