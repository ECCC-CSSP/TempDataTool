﻿//------------------------------------------------------------------------------
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
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    public partial class TempDataToolDBEntities : DbContext
    {
        public TempDataToolDBEntities()
            : base("name=TempDataToolDBEntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public virtual DbSet<ASGADMPNTable> ASGADMPNTables { get; set; }
        public virtual DbSet<ASGADPrecData> ASGADPrecDatas { get; set; }
        public virtual DbSet<ASGADPrecStn> ASGADPrecStns { get; set; }
        public virtual DbSet<ASGADRunMOU> ASGADRunMOUs { get; set; }
        public virtual DbSet<ASGADRun> ASGADRuns { get; set; }
        public virtual DbSet<ASGADSampCode> ASGADSampCodes { get; set; }
        public virtual DbSet<ASGADSampleMOU> ASGADSampleMOUs { get; set; }
        public virtual DbSet<ASGADSample> ASGADSamples { get; set; }
        public virtual DbSet<ASGADStation> ASGADStations { get; set; }
        public virtual DbSet<ASGADSubsect> ASGADSubsects { get; set; }
        public virtual DbSet<ASGADTide> ASGADTides { get; set; }
        public virtual DbSet<BCLandSample> BCLandSamples { get; set; }
        public virtual DbSet<BCLandSampleStation> BCLandSampleStations { get; set; }
        public virtual DbSet<BCMarineSample> BCMarineSamples { get; set; }
        public virtual DbSet<BCMarineSampleStation> BCMarineSampleStations { get; set; }
        public virtual DbSet<BCName> BCNames { get; set; }
        public virtual DbSet<BCPolSource> BCPolSources { get; set; }
        public virtual DbSet<BCPrecipitation> BCPrecipitations { get; set; }
        public virtual DbSet<BCSurvey> BCSurveys { get; set; }
        public virtual DbSet<LocationNameLatLng> LocationNameLatLngs { get; set; }
        public virtual DbSet<Nomenclature> Nomenclatures { get; set; }
        public virtual DbSet<QCClimateSite> QCClimateSites { get; set; }
        public virtual DbSet<QCSecteurMPol> QCSecteurMPols { get; set; }
        public virtual DbSet<QCSubsectorAssociation> QCSubsectorAssociations { get; set; }
        public virtual DbSet<QCSubSector> QCSubSectors { get; set; }
        public virtual DbSet<SanitaryDavCOb> SanitaryDavCObs { get; set; }
        public virtual DbSet<SanitaryDavCSite> SanitaryDavCSites { get; set; }
        public virtual DbSet<SanitaryDavOb> SanitaryDavObs { get; set; }
        public virtual DbSet<SanitaryDavSite> SanitaryDavSites { get; set; }
        public virtual DbSet<SanitaryDonOb> SanitaryDonObs { get; set; }
        public virtual DbSet<SanitaryDonSite> SanitaryDonSites { get; set; }
        public virtual DbSet<SanitaryJoeOb> SanitaryJoeObs { get; set; }
        public virtual DbSet<SanitaryJoeSite> SanitaryJoeSites { get; set; }
        public virtual DbSet<SanitaryPatOb> SanitaryPatObs { get; set; }
        public virtual DbSet<SanitaryPatSite> SanitaryPatSites { get; set; }
    }
}
