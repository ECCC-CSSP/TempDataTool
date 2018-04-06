using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TempDataTool
{
    public partial class TempDataTool : Form
    {
        #region Variables
        List<BackgroundWorker> bwListDon = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListPat = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListDav = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListJoe = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListDavC = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListSubsect = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListStation = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListSample = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListPrecData = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListRun = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListBCPolSource = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListBCPrec = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListBCRun = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListBCMarineStation = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListBCLandStation = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListBCMarineSample = new List<BackgroundWorker>();
        List<BackgroundWorker> bwListBCLandSample = new List<BackgroundWorker>();
        int DoBWCount = 200; // keep this number at 200
        string DonPatDav = "";
        bool PleaseStop = false;
        #endregion Variables

        #region Constructors
        public TempDataTool()
        {
            InitializeComponent();
            timerSanitary.Stop();
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkDon;

                bwListDon.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkPat;

                bwListPat.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkDav;

                bwListDav.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkJoe;

                bwListJoe.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkDavC;

                bwListDavC.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkSubsect;

                bwListSubsect.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkStation;

                bwListStation.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkSample;

                bwListSample.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkPrecData;

                bwListPrecData.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkRun;

                bwListRun.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkBCPolSource;

                bwListBCPolSource.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkBCPrec;

                bwListBCPrec.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkBCRun;

                bwListBCRun.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkBCMarineStation;

                bwListBCMarineStation.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkBCLandStation;

                bwListBCLandStation.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkBCMarineSample;

                bwListBCMarineSample.Add(bw);
            }
            for (int i = 0; i < DoBWCount; i++)
            {
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += bw_DoWorkBCLandSample;

                bwListBCLandSample.Add(bw);
            }
        }

        #endregion Constructors

        #region Events
        void bw_DoWorkBCLandSample(object sender, DoWorkEventArgs e)
        {
            BWBCLandSample bwobj = e.Argument as BWBCLandSample;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (BCLandSample bcPS in bwobj.bcLandSampleList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.BCLandSamples.Add(bcPS);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkBCLandStation(object sender, DoWorkEventArgs e)
        {
            BWBCLandStation bwobj = e.Argument as BWBCLandStation;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (BCLandSampleStation bcPS in bwobj.bcLandStationList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.BCLandSampleStations.Add(bcPS);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkBCMarineSample(object sender, DoWorkEventArgs e)
        {
            BWBCMarineSample bwobj = e.Argument as BWBCMarineSample;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (BCMarineSample bcPS in bwobj.bcMarineSampleList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.BCMarineSamples.Add(bcPS);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkBCMarineStation(object sender, DoWorkEventArgs e)
        {
            BWBCMarineStation bwobj = e.Argument as BWBCMarineStation;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (BCMarineSampleStation bcPS in bwobj.bcMarineStationList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.BCMarineSampleStations.Add(bcPS);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkBCPolSource(object sender, DoWorkEventArgs e)
        {
            BWBCPolSource bwobj = e.Argument as BWBCPolSource;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (BCPolSource bcPS in bwobj.bcPolSourceList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.BCPolSources.Add(bcPS);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkBCPrec(object sender, DoWorkEventArgs e)
        {
            BWBCPrec bwobj = e.Argument as BWBCPrec;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (BCPrecipitation bcPS in bwobj.bcPrecList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.BCPrecipitations.Add(bcPS);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkBCRun(object sender, DoWorkEventArgs e)
        {
            BWBCRun bwobj = e.Argument as BWBCRun;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (BCSurvey bcPS in bwobj.bcRunList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.BCSurveys.Add(bcPS);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkJoe(object sender, DoWorkEventArgs e)
        {
            BWJoe bwobj = e.Argument as BWJoe;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (SanitaryJoeSite ds in bwobj.siteList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.SanitaryJoeSites.Add(ds);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkDav(object sender, DoWorkEventArgs e)
        {
            BWDav bwobj = e.Argument as BWDav;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (SanitaryDavSite ds in bwobj.siteList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.SanitaryDavSites.Add(ds);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkDavC(object sender, DoWorkEventArgs e)
        {
            BWDavC bwDavC = e.Argument as BWDavC;
            DoDavCFunction(bwDavC);
        }

        void bw_DoWorkDon(object sender, DoWorkEventArgs e)
        {
            BWDon bwobj = e.Argument as BWDon;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (SanitaryDonSite ds in bwobj.siteList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.SanitaryDonSites.Add(ds);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkPat(object sender, DoWorkEventArgs e)
        {
            BWPat bwPat = e.Argument as BWPat;
            DoPatriceFunction(bwPat);
        }

        private void DoDavCFunction(BWDavC bwDavC)
        {
            List<SanitaryDavCSite> siteDavCList = new List<SanitaryDavCSite>();

            string connectionString = "";

            connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + bwDavC.FileName + ";Extended Properties=Excel 12.0";
            OleDbConnection conn = new OleDbConnection(connectionString);

            conn.Open();
            OleDbDataReader reader;

            Application.DoEvents();

            OleDbCommand comm = new OleDbCommand("Select * from [Sheet1$];");


            comm.Connection = conn;
            reader = comm.ExecuteReader();


            List<string> FieldNameList = new List<string>();
            FieldNameList = new List<string>() { "Locator", "Locator_Name", "Site", "Type", "Observations", "Zone", "Datum", "latitude", "longitude", "Active", "Date_Created" };
            for (int j = 0; j < reader.FieldCount; j++)
            {
                if (reader.GetName(j) != FieldNameList[j])
                {
                    richTextBox1.AppendText(bwDavC.FileName + " Sheet1 " + reader.GetName(j) + " is not equal to " + FieldNameList[j] + "\r\n");
                    return;
                }
            }
            reader.Close();

            reader = comm.ExecuteReader();


            int CountRead = 0;
            while (reader.Read())
            {
                CountRead += 1;
                //richTextBox1.AppendText(CountRead.ToString() + " ");

                Application.DoEvents();

                string Locator = "";
                //string Locator_Name = "";
                int Site = -1;
                string Type = "";
                string Observations = "";
                DateTime Date = new DateTime(2050, 1, 1);
                int Zone = -1;
                string Datum = "";
                float Latitude = 0.0f;
                float Longitude = 0.0f;
                //string Active = "";
                DateTime DateUpdated = new DateTime(2050, 1, 1);
                //string Photo = "";
                //string Civic = "";
                //string SampleMPN = "";
                string Status = "";

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    Locator = "";
                    continue;
                }
                else
                {
                    Locator = reader.GetValue(0).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    Site = -1;
                }
                else
                {
                    Site = int.Parse(reader.GetValue(2).ToString().Trim());
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    Type = "";
                }
                else
                {
                    Type = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    Observations = "";
                }
                else
                {
                    Observations = reader.GetValue(4).ToString().Trim();
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    Zone = -1;
                }
                else
                {
                    Zone = (int)(double)(reader.GetValue(5));
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    Datum = "";
                }
                else
                {
                    Datum = reader.GetValue(6).ToString().Trim();
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    Latitude = -1;
                }
                else
                {
                    Latitude = float.Parse(reader.GetValue(7).ToString().Trim());
                }

                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    Longitude = -1;
                }
                else
                {
                    Longitude = float.Parse(reader.GetValue(8).ToString().Trim());
                }

                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                {
                    DateUpdated = new DateTime(2050, 1, 1);
                }
                else
                {
                    DateUpdated = (DateTime)(reader.GetValue(10));
                }

                SanitaryDavCSite siteNew = new SanitaryDavCSite();
                siteNew.siteid = null;
                siteNew.Locator = Locator;
                siteNew.Site = Site;
                siteNew.Zone = Zone;
                siteNew.Datum = Datum;
                siteNew.latitude = Latitude;
                siteNew.longitude = Longitude;

                SanitaryDavCSite siteExist = siteDavCList.Where(c => c.Locator == Locator && c.Site == Site).FirstOrDefault<SanitaryDavCSite>();

                if (siteExist == null)
                {
                    SanitaryDavCOb obNew = new SanitaryDavCOb();
                    obNew.Active = true;
                    obNew.siteid = null;
                    obNew.ObsDate = DateUpdated;
                    obNew.Name_Inspector = "";
                    obNew.Type = Type;
                    obNew.Status = Status;
                    obNew.Risk_Assessment = "";
                    obNew.Observations = Observations;

                    siteNew.SanitaryDavCObs.Add(obNew);

                    siteDavCList.Add(siteNew);
                }
                else
                {
                    SanitaryDavCOb obNew = new SanitaryDavCOb();
                    obNew.Active = true;
                    obNew.siteid = null;
                    obNew.ObsDate = DateUpdated;
                    obNew.Name_Inspector = "";
                    obNew.Type = Type;
                    obNew.Status = Status;
                    obNew.Risk_Assessment = "";
                    obNew.Observations = Observations;

                    siteExist.SanitaryDavCObs.Add(obNew);
                }
            }
            reader.Close();
            //richTextBox1.AppendText("\r\n");


            conn.Close();

            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (SanitaryDavCSite sp in siteDavCList)
                {
                    dbDT.SanitaryDavCSites.Add(sp);
                    Application.DoEvents();

                }
                dbDT.SaveChanges();

            }
        }
        private void DoPatriceFunction(BWPat bwPat)
        {
            List<SanitaryPatSite> sitePatList = new List<SanitaryPatSite>();

            string connectionString = "";

            if (Path.GetExtension(bwPat.FileName) == ".xls")
            {
                connectionString = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + bwPat.FileName + @";Extended Properties=Excel 8.0";
            }
            else if (Path.GetExtension(bwPat.FileName) == ".xlsx")
            {
                connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + bwPat.FileName + ";Extended Properties=Excel 12.0";
            }
            OleDbConnection conn = new OleDbConnection(connectionString);

            conn.Open();
            OleDbDataReader reader;

            for (int i = 1995; i < 2014; i++)
            {
                Application.DoEvents();

                OleDbCommand comm = new OleDbCommand("Select * from [" + i + "$];");

                try
                {
                    comm.Connection = conn;
                    reader = comm.ExecuteReader();
                }
                catch (Exception) // not a real error just trying to find the year of the sheet
                {
                    continue;
                }

                List<string> FieldNameList = new List<string>();
                FieldNameList = new List<string>() { "Locator", "Locator_Name", "Site", "Type", "Observations", "Date", "Zone", "Datum", "Easting", "Northing", "Latitude", "Longitude", "PS or NP", "Date Updated", "Photo", "Civic #", "Sample (MPN/100ml)", "Status" };
                for (int j = 0; j < reader.FieldCount; j++)
                {
                    if (reader.GetName(j) != FieldNameList[j])
                    {
                        richTextBox1.AppendText(bwPat.FileName + " sheet " + i + reader.GetName(j) + " is not equal to " + FieldNameList[j] + "\r\n");
                        return;
                    }
                }
                reader.Close();

                reader = comm.ExecuteReader();


                int CountRead = 0;
                while (reader.Read())
                {
                    CountRead += 1;
                    //richTextBox1.AppendText(CountRead.ToString() + " ");

                    Application.DoEvents();

                    string Locator = "";
                    //string Locator_Name = "";
                    int Site = -1;
                    string Type = "";
                    string Observations = "";
                    DateTime Date = new DateTime(2050, 1, 1);
                    int Zone = -1;
                    string Datum = "";
                    float Easting = 0f;
                    float Northing = 0f;
                    float Latitude = 0.0f;
                    float Longitude = 0.0f;
                    //string PSOrNP = "";
                    DateTime DateUpdated = new DateTime(2050, 1, 1);
                    //string Photo = "";
                    //string Civic = "";
                    //string SampleMPN = "";
                    string Status = "";

                    if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                    {
                        Locator = "";
                        continue;
                    }
                    else
                    {
                        Locator = reader.GetValue(0).ToString().Trim();
                    }

                    if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                    {
                        Site = -1;
                    }
                    else
                    {
                        Site = int.Parse(reader.GetValue(2).ToString().Trim());
                    }

                    if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                    {
                        Type = "";
                    }
                    else
                    {
                        Type = reader.GetValue(3).ToString().Trim();
                    }

                    if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                    {
                        Observations = "";
                    }
                    else
                    {
                        Observations = reader.GetValue(4).ToString().Trim();
                    }

                    if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                    {
                        Zone = -1;
                    }
                    else
                    {
                        Zone = (int)(double)(reader.GetValue(6));
                    }

                    if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                    {
                        Datum = "";
                    }
                    else
                    {
                        Datum = reader.GetValue(7).ToString().Trim();
                    }

                    if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                    {
                        Easting = -1;
                    }
                    else
                    {
                        Easting = float.Parse(reader.GetValue(8).ToString().Trim());
                    }

                    if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                    {
                        Northing = -1;
                    }
                    else
                    {
                        Northing = float.Parse(reader.GetValue(9).ToString().Trim());
                    }

                    if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                    {
                        Latitude = -1;
                    }
                    else
                    {
                        Latitude = float.Parse(reader.GetValue(10).ToString().Trim());
                    }

                    if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                    {
                        Longitude = -1;
                    }
                    else
                    {
                        Longitude = float.Parse(reader.GetValue(11).ToString().Trim());
                    }

                    if (reader.GetValue(13).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(13).ToString()))
                    {
                        DateUpdated = new DateTime(2050, 1, 1);
                    }
                    else
                    {
                        DateUpdated = (DateTime)(reader.GetValue(13));
                    }

                    if (reader.GetValue(17).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(17).ToString()))
                    {
                        Status = "";
                    }
                    else
                    {
                        Status = reader.GetValue(17).ToString().Trim();
                    }

                    SanitaryPatSite siteNew = new SanitaryPatSite();
                    siteNew.siteid = null;
                    siteNew.Locator = Locator;
                    siteNew.Site = Site;
                    siteNew.Zone = Zone;
                    siteNew.easting = Easting;
                    siteNew.northing = Northing;
                    siteNew.Datum = Datum;
                    siteNew.latitude = Latitude;
                    siteNew.longitude = Longitude;

                    SanitaryPatSite siteExist = sitePatList.Where(c => c.Locator == Locator && c.Site == Site).FirstOrDefault<SanitaryPatSite>();

                    if (siteExist == null)
                    {
                        SanitaryPatOb obNew = new SanitaryPatOb();
                        obNew.Active = true;
                        obNew.siteid = null;
                        obNew.ObsDate = DateUpdated;
                        obNew.Name_Inspector = "";
                        obNew.Type = Type;
                        obNew.Status = Status;
                        obNew.Risk_Assessment = "";
                        obNew.Observations = Observations;

                        siteNew.SanitaryPatObs.Add(obNew);

                        sitePatList.Add(siteNew);
                    }
                    else
                    {
                        SanitaryPatOb obNew = new SanitaryPatOb();
                        obNew.Active = true;
                        obNew.siteid = null;
                        obNew.ObsDate = DateUpdated;
                        obNew.Name_Inspector = "";
                        obNew.Type = Type;
                        obNew.Status = Status;
                        obNew.Risk_Assessment = "";
                        obNew.Observations = Observations;

                        siteExist.SanitaryPatObs.Add(obNew);
                    }
                }
                reader.Close();
                //richTextBox1.AppendText("\r\n");

            }

            conn.Close();

            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (SanitaryPatSite sp in sitePatList)
                {
                    dbDT.SanitaryPatSites.Add(sp);
                    Application.DoEvents();
                }

                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkRun(object sender, DoWorkEventArgs e)
        {
            BWRun bwobj = e.Argument as BWRun;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (ASGADRun apd in bwobj.asgadRunList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.ASGADRuns.Add(apd);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkPrecData(object sender, DoWorkEventArgs e)
        {
            BWPrecData bwobj = e.Argument as BWPrecData;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (ASGADPrecData apd in bwobj.asgadPrecDataList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.ASGADPrecDatas.Add(apd);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkSample(object sender, DoWorkEventArgs e)
        {
            BWSample bwobj = e.Argument as BWSample;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (ASGADSample asa in bwobj.asgadSampleList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.ASGADSamples.Add(asa);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkStation(object sender, DoWorkEventArgs e)
        {
            BWStation bwobj = e.Argument as BWStation;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (ASGADStation ast in bwobj.asgadStationList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.ASGADStations.Add(ast);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        void bw_DoWorkSubsect(object sender, DoWorkEventArgs e)
        {
            BWSubsect bwobj = e.Argument as BWSubsect;
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (ASGADSubsect ass in bwobj.asgadSubsectList.Skip(bwobj.skip).Take(bwobj.take))
                {
                    dbDT.ASGADSubsects.Add(ass);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        private void butFixMissingStations_Click(object sender, EventArgs e)
        {
            FixMissingStations();
        }
        private void butLoadASGAD_Click(object sender, EventArgs e)
        {
            //richTextBox1.Text = DateTime.Now + "\r\n";
            ////
            //// ---- doing ASGADSamples
            ////
            //CleanASGADSamples("ZZ");
            //DonPatDav = "Sample";
            //PleaseStop = false;
            //timerSanitary.Start();
            ////LoadASGADSample("NB");
            ////bool BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListSample)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}

            ////lblStatus.Text = "LoadASGADSample NB done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();

            //LoadASGADSample("NL");
            //bool BWBusy = true;
            //while (BWBusy)
            //{
            //    Application.DoEvents();
            //    BWBusy = false;
            //    foreach (BackgroundWorker bw in bwListSample)
            //    {
            //        if (bw.IsBusy)
            //        {
            //            BWBusy = true;
            //            break;
            //        }
            //    }
            //}

            //lblStatus.Text = "LoadASGADSample NL done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            ////LoadASGADSample("NS");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListSample)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}

            ////lblStatus.Text = "LoadASGADSample NS done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();

            ////LoadASGADSample("PE");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListSample)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}

            ////lblStatus.Text = "LoadASGADSample PE done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();

            ////
            //// ---- doing ASGADRuns
            ////
            //CleanASGADRuns();
            //DonPatDav = "Run";
            //PleaseStop = false;
            //timerSanitary.Start();
            ////LoadASGADRun("NB");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListRun)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}

            ////lblStatus.Text = "LoadASGADRuns NB done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();

            //LoadASGADRun("NL");
            //BWBusy = true;
            //while (BWBusy)
            //{
            //    Application.DoEvents();
            //    BWBusy = false;
            //    foreach (BackgroundWorker bw in bwListRun)
            //    {
            //        if (bw.IsBusy)
            //        {
            //            BWBusy = true;
            //            break;
            //        }
            //    }
            //}

            //lblStatus.Text = "LoadASGADRuns NL done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            ////LoadASGADRun("NS");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListRun)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}

            ////lblStatus.Text = "LoadASGADRuns NS done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();

            ////LoadASGADRun("PE");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListRun)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}


            ////lblStatus.Text = "LoadASGADRuns PE done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();


            //LoadASGADMPNTABLE("NB");

            ////
            //// ---- doing ASGADPrecDatas
            ////
            //CleanASGADPrecDatas();
            //DonPatDav = "PrecData";
            //PleaseStop = false;
            //timerSanitary.Start();
            ////LoadASGADPrecData("NB");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListPrecData)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}

            ////lblStatus.Text = "LoadASGADPrecData NB done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();

            //LoadASGADPrecData("NL");
            //BWBusy = true;
            //while (BWBusy)
            //{
            //    Application.DoEvents();
            //    BWBusy = false;
            //    foreach (BackgroundWorker bw in bwListPrecData)
            //    {
            //        if (bw.IsBusy)
            //        {
            //            BWBusy = true;
            //            break;
            //        }
            //    }
            //}

            //lblStatus.Text = "LoadASGADPrecData NL done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            ////LoadASGADPrecData("NS");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListPrecData)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}

            ////lblStatus.Text = "LoadASGADPrecData NS done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();

            ////LoadASGADPrecData("PE");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListPrecData)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}


            ////lblStatus.Text = "LoadASGADPrecData PE done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();


            ////CleanASGADPrecStns();
            ////LoadASGADPrecStn("NB");


            ////lblStatus.Text = "LoadASGADPrecStn NB done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();


            ////CleanASGADSampCodes();
            ////LoadASGADSampCode("NB");


            ////lblStatus.Text = "LoadASGADSampCode NB done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();


            ////
            //// ---- doing ASGADStations
            ////
            //CleanASGADStations();
            //DonPatDav = "Station";
            //PleaseStop = false;
            //timerSanitary.Start();
            ////LoadASGADStation("NB");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListStation)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}

            ////lblStatus.Text = "LoadASGADStation NB done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();

            //LoadASGADStation("NL");
            //BWBusy = true;
            //while (BWBusy)
            //{
            //    Application.DoEvents();
            //    BWBusy = false;
            //    foreach (BackgroundWorker bw in bwListStation)
            //    {
            //        if (bw.IsBusy)
            //        {
            //            BWBusy = true;
            //            break;
            //        }
            //    }
            //}

            //lblStatus.Text = "LoadASGADStation NL done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            ////LoadASGADStation("NS");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListStation)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}

            ////lblStatus.Text = "LoadASGADStation NS done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();

            ////LoadASGADStation("PE");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListStation)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}

            ////lblStatus.Text = "LoadASGADStation PE done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();

            //// 
            //// -------- doing subsect
            ////

            //CleanASGADSubsect();
            //DonPatDav = "Subsect";
            //PleaseStop = false;
            //timerSanitary.Start();
            ////LoadASGADSubsect("NB");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListSubsect)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}

            ////lblStatus.Text = "LoadASGADSubsect NB done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();

            //LoadASGADSubsect("NL");
            //BWBusy = true;
            //while (BWBusy)
            //{
            //    Application.DoEvents();
            //    BWBusy = false;
            //    foreach (BackgroundWorker bw in bwListSubsect)
            //    {
            //        if (bw.IsBusy)
            //        {
            //            BWBusy = true;
            //            break;
            //        }
            //    }
            //}

            //lblStatus.Text = "LoadASGADSubsect NL done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            ////LoadASGADSubsect("NS");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListSubsect)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}

            ////lblStatus.Text = "LoadASGADSubsect NS done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();

            ////LoadASGADSubsect("PE");
            ////BWBusy = true;
            ////while (BWBusy)
            ////{
            ////    Application.DoEvents();
            ////    BWBusy = false;
            ////    foreach (BackgroundWorker bw in bwListSubsect)
            ////    {
            ////        if (bw.IsBusy)
            ////        {
            ////            BWBusy = true;
            ////            break;
            ////        }
            ////    }
            ////}

            ////lblStatus.Text = "LoadASGADSubsect PE done ...";
            ////lblStatus.Refresh();
            ////Application.DoEvents();

            //CleanASGADTides();
            //LoadASGADTide("NB");

            //using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            //{
            //    List<ASGADSubsect> asgadSubsecList = (from c in dbDT.ASGADSubsects
            //                                          where c.PRECIPID == "8202534"
            //                                          select c).ToList<ASGADSubsect>();

            //    foreach (ASGADSubsect ass in asgadSubsecList)
            //    {
            //        ass.PRECIPID = "8202535";
            //    }

            //    dbDT.SaveChanges();
            //}

            //richTextBox1.AppendText(DateTime.Now + "\r\n");
        }
        private void butLoadASGADNB_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = DateTime.Now + "\r\n";

            //
            // ---- doing ASGADStations
            //
            lblStatus.Text = "Cleaning NB Stations  ...";
            lblStatus.Refresh();
            Application.DoEvents();
            CleanASGADStations("NB");
            lblStatus.Text = "Cleaning NB Stations done  ...";
            lblStatus.Refresh();
            Application.DoEvents();

            lblStatus.Text = "LoadASGADStation NB  ...";
            lblStatus.Refresh();
            Application.DoEvents();
            LoadASGADStation("NB");
            lblStatus.Text = "LoadASGADStation NB done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            //
            // ---- doing ASGADRuns
            //
            lblStatus.Text = "Cleaning NB Runs  ...";
            lblStatus.Refresh();
            Application.DoEvents();
            CleanASGADRuns("NB");
            lblStatus.Text = "Cleaning NB Runs done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            lblStatus.Text = "LoadASGADRuns NB  ...";
            lblStatus.Refresh();
            Application.DoEvents();
            LoadASGADRun("NB");
            lblStatus.Text = "LoadASGADRuns NB done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            ////
            //// ---- doing ASGADPrecDatas
            ////
            //lblStatus.Text = "Cleaning NB Prec data  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //CleanASGADPrecDatas("NB");
            //lblStatus.Text = "Cleaning NB Prec data done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //lblStatus.Text = "LoadASGADPrecData NB  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //LoadASGADPrecData("NB");
            //lblStatus.Text = "LoadASGADPrecData NB done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();


            //
            // ---- doing ASGAD NBSamples
            //
            lblStatus.Text = "Cleaning NB Samples ...";
            lblStatus.Refresh();
            Application.DoEvents();
            CleanASGADSamples("NB");
            lblStatus.Text = "Cleaning NB Samples Done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            lblStatus.Text = "LoadASGADSample NB ...";
            lblStatus.Refresh();
            Application.DoEvents();
            LoadASGADSample("NB");
            lblStatus.Text = "LoadASGADSample NB done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            richTextBox1.AppendText(DateTime.Now + "\r\n");
        }
        private void butLoadASGADNBMOU_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = DateTime.Now + "\r\n";

            ////
            //// ---- doing ASGADStations
            ////
            //lblStatus.Text = "Cleaning NB Stations  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //CleanASGADStations("NB");
            //lblStatus.Text = "Cleaning NB Stations done  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //lblStatus.Text = "LoadASGADStation NB  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //LoadASGADStation("NB");
            //lblStatus.Text = "LoadASGADStation NB done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //
            // ---- doing ASGADRuns
            //
            lblStatus.Text = "Cleaning NB Runs  ...";
            lblStatus.Refresh();
            Application.DoEvents();
            CleanASGADRunMOUs("NB");
            lblStatus.Text = "Cleaning NB Runs done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            lblStatus.Text = "LoadASGADRuns NB  ...";
            lblStatus.Refresh();
            Application.DoEvents();
            LoadASGADRunMOU("NB");
            lblStatus.Text = "LoadASGADRuns NB done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            //
            // ---- doing ASGAD NBSamples
            //
            lblStatus.Text = "Cleaning NB Samples ...";
            lblStatus.Refresh();
            Application.DoEvents();
            CleanASGADSampleMOUs("NB");
            lblStatus.Text = "Cleaning NB Samples Done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            lblStatus.Text = "LoadASGADSample NB ...";
            lblStatus.Refresh();
            Application.DoEvents();
            LoadASGADSampleMOU("NB");
            lblStatus.Text = "LoadASGADSample NB done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            richTextBox1.AppendText(DateTime.Now + "\r\n");
        }
        private void butLoadASGADNL_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = DateTime.Now + "\r\n";

            //
            // ---- doing ASGADStations
            //
            lblStatus.Text = "Cleaning NL Stations  ...";
            lblStatus.Refresh();
            Application.DoEvents();
            CleanASGADStations("NL");
            lblStatus.Text = "Cleaning NL Stations done  ...";
            lblStatus.Refresh();
            Application.DoEvents();

            lblStatus.Text = "LoadASGADStation NL  ...";
            lblStatus.Refresh();
            Application.DoEvents();
            LoadASGADStation("NL");
            lblStatus.Text = "LoadASGADStation NL done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            //
            // ---- doing ASGADRuns
            //
            lblStatus.Text = "Cleaning NL Runs  ...";
            lblStatus.Refresh();
            Application.DoEvents();
            CleanASGADRuns("NL");
            lblStatus.Text = "Cleaning NL Runs done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            lblStatus.Text = "LoadASGADRuns NL  ...";
            lblStatus.Refresh();
            Application.DoEvents();
            LoadASGADRun("NL");
            lblStatus.Text = "LoadASGADRuns NL done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            ////
            //// ---- doing ASGADPrecDatas
            ////
            //lblStatus.Text = "Cleaning NL Prec data  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //CleanASGADPrecDatas("NL");
            //lblStatus.Text = "Cleaning NL Prec data done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //lblStatus.Text = "LoadASGADPrecData NL  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //LoadASGADPrecData("NL");
            //lblStatus.Text = "LoadASGADPrecData NL done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();


            //
            // ---- doing ASGAD NLSamples
            //
            lblStatus.Text = "Cleaning NL Samples ...";
            lblStatus.Refresh();
            Application.DoEvents();
            CleanASGADSamples("NL");
            lblStatus.Text = "Cleaning NL Samples Done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            lblStatus.Text = "LoadASGADSample NL ...";
            lblStatus.Refresh();
            Application.DoEvents();
            LoadASGADSample("NL");
            lblStatus.Text = "LoadASGADSample NL done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            richTextBox1.AppendText(DateTime.Now + "\r\n");
        }
        private void butLoadASGADNS_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = DateTime.Now + "\r\n";

            //
            // ---- doing ASGADStations
            //
            //lblStatus.Text = "Cleaning NS Stations  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //CleanASGADStations("NS");
            //lblStatus.Text = "Cleaning NS Stations done  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //lblStatus.Text = "LoadASGADStation NS  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //LoadASGADStation("NS");
            //lblStatus.Text = "LoadASGADStation NS done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //
            // ---- doing ASGADRuns
            //
            //lblStatus.Text = "Cleaning NS Runs  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //CleanASGADRuns("NS");
            //lblStatus.Text = "Cleaning NS Runs done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //lblStatus.Text = "LoadASGADRuns NS  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //LoadASGADRun("NS");
            //lblStatus.Text = "LoadASGADRuns NS done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //
            // ---- doing ASGADPrecDatas
            //
            //lblStatus.Text = "Cleaning NS Prec data  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //CleanASGADPrecDatas("NS");
            //lblStatus.Text = "Cleaning NS Prec data done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //lblStatus.Text = "LoadASGADPrecData NS  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //LoadASGADPrecData("NS");
            //lblStatus.Text = "LoadASGADPrecData NS done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();


            //
            // ---- doing ASGAD NSSamples
            //
            //lblStatus.Text = "Cleaning NS Samples ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //CleanASGADSamples("NS");
            //lblStatus.Text = "Cleaning NS Samples Done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //lblStatus.Text = "LoadASGADSample NS ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //LoadASGADSample("NS");
            //lblStatus.Text = "LoadASGADSample NS done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //richTextBox1.AppendText(DateTime.Now + "\r\n");

        }
        private void butLoadASGADPE_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = DateTime.Now + "\r\n";

            //
            // ---- doing ASGADStations
            //
            //lblStatus.Text = "Cleaning PE Stations  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //CleanASGADStations("PE");
            //lblStatus.Text = "Cleaning PE Stations done  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //lblStatus.Text = "LoadASGADStation PE  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //LoadASGADStation("PE");
            //lblStatus.Text = "LoadASGADStation PE done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //
            // ---- doing ASGADRuns
            //
            //lblStatus.Text = "Cleaning PE Runs  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //CleanASGADRuns("PE");
            //lblStatus.Text = "Cleaning PE Runs done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //lblStatus.Text = "LoadASGADRuns PE  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //LoadASGADRun("PE");
            //lblStatus.Text = "LoadASGADRuns PE done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //
            // ---- doing ASGADPrecDatas
            //
            //lblStatus.Text = "Cleaning PE Prec data  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //CleanASGADPrecDatas("PE");
            //lblStatus.Text = "Cleaning PE Prec data done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            //lblStatus.Text = "LoadASGADPrecData PE  ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //LoadASGADPrecData("PE");
            //lblStatus.Text = "LoadASGADPrecData PE done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();


            //
            // ---- doing ASGAD PESamples
            //
            //lblStatus.Text = "Cleaning PE Samples ...";
            //lblStatus.Refresh();
            //Application.DoEvents();
            //CleanASGADSamples("PE");
            //lblStatus.Text = "Cleaning PE Samples Done ...";
            //lblStatus.Refresh();
            //Application.DoEvents();

            lblStatus.Text = "LoadASGADSample PE ...";
            lblStatus.Refresh();
            Application.DoEvents();
            LoadASGADSample("PE");
            lblStatus.Text = "LoadASGADSample PE done ...";
            lblStatus.Refresh();
            Application.DoEvents();

            richTextBox1.AppendText(DateTime.Now + "\r\n");

        }
        private void butLoadBCLand_Base_Sample_Readings_Click(object sender, EventArgs e)
        {
            DonPatDav = "BCLand_Base_Sample_Readings";
            PleaseStop = false;
            timerSanitary.Start();
            CleanBCLandSamples();
            LoadBCLand_Base_Sample_Readings();
            bool BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListBCLandSample)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }
        }
        private void butLoadBCLand_Based_Sample_Stations_Click(object sender, EventArgs e)
        {
            DonPatDav = "BCLand_Based_Sample_Stations";
            PleaseStop = false;
            timerSanitary.Start();
            CleanBCLandSampleStations();
            LoadBCLand_Based_Sample_Stations();
            bool BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListBCLandStation)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }
        }
        private void butLoadBCPolSource_Click(object sender, EventArgs e)
        {
            DonPatDav = "BCPolSource";
            PleaseStop = false;
            timerSanitary.Start();
            CleanBCPolSources();
            LoadBCPolSource();
            bool BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListBCPolSource)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }
        }
        private void butLoadBCMarine_Sample_Readings_Click(object sender, EventArgs e)
        {
            DonPatDav = "BCMarine_Sample_Readings";
            PleaseStop = false;
            timerSanitary.Start();
            CleanBCMarineSamples();
            LoadBCMarine_Sample_Readings();
            bool BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListBCMarineSample)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }
        }
        private void butLoadBCMarine_Sample_Stations_Click(object sender, EventArgs e)
        {
            DonPatDav = "BCMarine_Sample_Stations";
            PleaseStop = false;
            timerSanitary.Start();
            CleanBCMarineSampleStations();
            LoadBCMarine_Sample_Stations();
            bool BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListBCMarineStation)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }
        }
        private void butLoadBCSurveys_Click(object sender, EventArgs e)
        {
            DonPatDav = "BCSurveys";
            PleaseStop = false;
            timerSanitary.Start();
            CleanBCSurveys();
            LoadBCSurveys();
            bool BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListBCRun)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }
        }
        private void butLoadBCWeather_Readings_Click(object sender, EventArgs e)
        {
            DonPatDav = "BCWeather_Readings";
            PleaseStop = false;
            timerSanitary.Start();
            CleanBCPrecipitations();
            LoadBCWeather_Readings();
            bool BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListBCPrec)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }
        }
        private void butLoadBCWeatherMissing_Click(object sender, EventArgs e)
        {
            DonPatDav = "BCWeatherMissing";
            PleaseStop = false;
            timerSanitary.Start();
            LoadBCWeatherMissing();
        }
        private void butLoadNomenclature_Click(object sender, EventArgs e)
        {
            LoadNomenclature();
        }
        private void butLoadSanitary_Click(object sender, EventArgs e)
        {
            DonPatDav = "Don";
            PleaseStop = false;
            timerSanitary.Start();
            CleanSanitaryDon();
            LoadSanitaryDon("NB");
            bool BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListDon)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }
            LoadSanitaryDon("NS");
            BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListDon)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }
            LoadSanitaryDon("PEI");
            BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListDon)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }

            butLoadSanitary.Text = "Load Sanitary Done...";
            lblStatus.Text = "Done ...";
            PleaseStop = true;
        }
        private void butLoadSanitaryPatrice_Click(object sender, EventArgs e)
        {
            DonPatDav = "Pat";
            lblTotal.Text = "2304";
            PleaseStop = false;
            timerSanitary.Start();
            CleanSanitaryPatrice();
            LoadSanitaryPatrice();
            bool BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListPat)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }
            butLoadSanitaryPatrice.Text = "Load Sanitary Patrice Done ...";
            lblStatus.Text = "Done ... ";
            PleaseStop = true;
        }
        private void butLoadSanitaryDave_Click(object sender, EventArgs e)
        {
            DonPatDav = "Dav";
            timerSanitary.Start();
            PleaseStop = false;
            CleanSanitaryDav();
            LoadSanitaryDav("NS");
            bool BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListDav)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }
            LoadSanitaryDav("PEI");
            BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListDav)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }

            butLoadSanitaryDave.Text = "Load Sanitary Dave Done...";
            lblStatus.Text = "Done ...";
            PleaseStop = true;
        }
        private void butLoadSanitaryJoe_Click(object sender, EventArgs e)
        {
            DonPatDav = "Joe";
            timerSanitary.Start();
            PleaseStop = false;
            CleanSanitaryJoe();
            LoadSanitaryJoe("NB");
            bool BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListDav)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }

            butLoadSanitaryJoe.Text = "Load Sanitary Joe Done...";
            lblStatus.Text = "Done ...";
            PleaseStop = true;
        }
        private void butLoadSanitaryDaveC_Click(object sender, EventArgs e)
        {
            PleaseStop = false;
            timerSanitary.Start();
            CleanSanitaryDaveC();
            LoadSanitaryDaveC();
            bool BWBusy = true;
            while (BWBusy)
            {
                Application.DoEvents();
                BWBusy = false;
                foreach (BackgroundWorker bw in bwListDavC)
                {
                    if (bw.IsBusy)
                    {
                        BWBusy = true;
                        break;
                    }
                }
            }
            butLoadSanitaryDaveC.Text = "Load Sanitary Dave C Done ...";
            lblStatus.Text = "Done ... ";
            PleaseStop = true;
        }
        private void timerSanitary_Tick(object sender, EventArgs e)
        {
            using (TempDataToolDBEntities dbTD = new TempDataToolDBEntities())
            {
                switch (DonPatDav)
                {
                    case "Don":
                        {
                            lblCountInDB.Text = (from c in dbTD.SanitaryDonSites
                                                 select c).Count().ToString();
                        }
                        break;
                    case "Pat":
                        {
                            lblCountInDB.Text = (from c in dbTD.SanitaryPatSites
                                                 select c).Count().ToString();
                        }
                        break;
                    case "Dav":
                        {
                            lblCountInDB.Text = (from c in dbTD.SanitaryDavSites
                                                 select c).Count().ToString();
                        }
                        break;
                    case "Subsect":
                        {
                            lblCountInDB.Text = (from c in dbTD.ASGADSubsects
                                                 select c).Count().ToString();
                        }
                        break;
                    case "Station":
                        {
                            lblCountInDB.Text = (from c in dbTD.ASGADStations
                                                 select c).Count().ToString();
                        }
                        break;
                    case "Sample":
                        {
                            lblCountInDB.Text = (from c in dbTD.ASGADSamples
                                                 select c).Count().ToString();
                        }
                        break;
                    case "Run":
                        {
                            lblCountInDB.Text = (from c in dbTD.ASGADRuns
                                                 select c).Count().ToString();
                        }
                        break;
                    case "PrecData":
                        {
                            lblCountInDB.Text = (from c in dbTD.ASGADPrecDatas
                                                 select c).Count().ToString();
                        }
                        break;
                    case "BCLand_Base_Sample_Readings":
                        {
                            lblCountInDB.Text = (from c in dbTD.BCLandSamples
                                                 select c).Count().ToString();
                        }
                        break;
                    case "BCLand_Based_Sample_Stations":
                        {
                            lblCountInDB.Text = (from c in dbTD.BCLandSampleStations
                                                 select c).Count().ToString();
                        }
                        break;
                    case "BCMarine_Sample_Readings":
                        {
                            lblCountInDB.Text = (from c in dbTD.BCMarineSamples
                                                 select c).Count().ToString();
                        }
                        break;
                    case "BCMarine_Sample_Stations":
                        {
                            lblCountInDB.Text = (from c in dbTD.BCMarineSampleStations
                                                 select c).Count().ToString();
                        }
                        break;
                    case "BCPolSource":
                        {
                            lblCountInDB.Text = (from c in dbTD.BCPolSources
                                                 select c).Count().ToString();
                        }
                        break;
                    case "BCSurveys":
                        {
                            lblCountInDB.Text = (from c in dbTD.BCSurveys
                                                 select c).Count().ToString();
                        }
                        break;
                    case "BCWeather_Readings":
                        {
                            lblCountInDB.Text = (from c in dbTD.BCPrecipitations
                                                 select c).Count().ToString();
                        }
                        break;
                    case "BCWeatherMissing":
                        {
                            int CountLeftToDo = (from c in dbTD.BCPrecipitations
                                                 where c.ClimateID == null
                                                 select c).Count();

                            lblCountInDB.Text = (int.Parse(lblTotal.Text) - CountLeftToDo).ToString();
                        }
                        break;
                    default:
                        break;
                }

                lblCountInDB.Refresh();
                Application.DoEvents();
            }

            if (PleaseStop)
            {
                timerSanitary.Stop();
            }
        }

        #endregion Events

        #region Functions
        private void CleanASGADMPNTables()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM ASGADMPNTables");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('ASGADMPNTables', RESEED, 0)");
            }
        }
        private void CleanASGADPrecDatas(string prov)
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM ASGADPrecDatas Where p = '" + prov + "'");
                //dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('ASGADPrecDatas', RESEED, 0)");
            }
        }
        private void CleanASGADPrecStns()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM ASGADPrecStns");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('ASGADPrecStns', RESEED, 0)");
            }
        }
        private void CleanASGADRuns(string prov)
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM ASGADRuns Where prov = '" + prov + "'");
                //dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('ASGADRuns', RESEED, 0)");
            }
        }
        private void CleanASGADRunMOUs(string prov)
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM ASGADRunMOUs Where prov = '" + prov + "'");
                //dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('ASGADRuns', RESEED, 0)");
            }
        }
        private void CleanASGADSampCodes()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM ASGADSampCodes");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('ASGADSampCodes', RESEED, 0)");
            }
        }
        private void CleanASGADSamples(string prov)
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM ASGADSamples Where Prov = '" + prov + "'");
                //dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('ASGADSamples', RESEED, 0)");
            }
        }
        private void CleanASGADSampleMOUs(string prov)
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM ASGADSampleMOUs Where Prov = '" + prov + "'");
                //dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('ASGADSamples', RESEED, 0)");
            }
        }
        private void CleanASGADStations(string prov)
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM ASGADStations Where Prov = '" + prov + "'");
                //dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('ASGADStations', RESEED, 0)");
            }
        }
        private void CleanASGADSubsect()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM ASGADSubsects");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('ASGADSubsects', RESEED, 0)");
            }
        }
        private void CleanASGADTides(string prov)
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM ASGADTides Where Prov = '" + prov + "'");
                //dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('ASGADTides', RESEED, 0)");
            }
        }
        private void CleanBCPolSources()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM BCPolSources");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('BCPolSources', RESEED, 0)");
            }
        }
        private void CleanBCLandSamples()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM BCLandSamples");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('BCLandSamples', RESEED, 0)");
            }
        }
        private void CleanBCLandSampleStations()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM BCLandSampleStations");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('BCLandSampleStations', RESEED, 0)");
            }
        }
        private void CleanBCMarineSamples()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM BCMarineSamples");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('BCMarineSamples', RESEED, 0)");
            }
        }
        private void CleanBCMarineSampleStations()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM BCMarineSampleStations");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('BCMarineSampleStations', RESEED, 0)");
            }
        }
        private void CleanBCSurveys()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM BCSurveys");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('BCSurveys', RESEED, 0)");
            }
        }
        private void CleanBCPrecipitations()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM BCPrecipitations");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('BCPrecipitations', RESEED, 0)");
            }
        }
        private void CleanSanitaryDav()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM SanitaryDavObs");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('SanitaryDavObs', RESEED, 0)");
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM SanitaryDavSites");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('SanitaryDavSites', RESEED, 0)");
            }
        }
        private void CleanSanitaryJoe()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM SanitaryJoeObs");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('SanitaryJoeObs', RESEED, 0)");
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM SanitaryJoeSites");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('SanitaryJoeSites', RESEED, 0)");
            }
        }
        private void CleanSanitaryDon()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM SanitaryDonObs");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('SanitaryDonObs', RESEED, 0)");
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM SanitaryDonSites");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('SanitaryDonSites', RESEED, 0)");
            }
        }
        private void CleanSanitaryPatrice()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM SanitaryPatObs");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('SanitaryPatObs', RESEED, 0)");
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM SanitaryPatSites");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('SanitaryPatSites', RESEED, 0)");
            }
        }
        private void CleanSanitaryDaveC()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM SanitaryDavCObs");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('SanitaryDavCObs', RESEED, 0)");
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM SanitaryDavCSites");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('SanitaryDavCSites', RESEED, 0)");
            }
        }
        private void FixMissingStations()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                List<StnInfo> StnInfoList = (from c in dbDT.ASGADSamples
                                             orderby c.p, c.AREA, c.SECTOR, c.SUBSECTOR, c.STAT_NBR
                                             select new StnInfo
                                             {
                                                 Prov = c.p,
                                                 Area = c.AREA,
                                                 Sector = c.SECTOR,
                                                 Subsector = c.SUBSECTOR,
                                                 Stn_Nbr = c.STAT_NBR,
                                             }).Distinct().ToList<StnInfo>();


                int TotalCount = StnInfoList.Count();
                int Count = 0;
                int orderNumb = 1000;
                foreach (StnInfo si in StnInfoList)
                {
                    Count += 1;
                    lblStatus.Text = Count + " of " + TotalCount;
                    lblStatus.Refresh();
                    Application.DoEvents();

                    ASGADStation asgadStationExist = (from c in dbDT.ASGADStations
                                                      where c.PROV == si.Prov
                                                      && c.AREA == si.Area
                                                      && c.SECTOR == si.Sector
                                                      && c.SUBSECTOR == si.Subsector
                                                      && c.STAT_NBR == si.Stn_Nbr
                                                      orderby c.AREA, c.SECTOR, c.SUBSECTOR, c.STAT_NBR
                                                      select c).FirstOrDefault<ASGADStation>();

                    if (asgadStationExist == null)
                    {
                        orderNumb += 1;
                        ASGADStation asgadStaNew = new ASGADStation();
                        asgadStaNew.ACTIVE = null;
                        asgadStaNew.AREA = si.Area;
                        asgadStaNew.EASTING = "";
                        asgadStaNew.LAT = 0f;
                        asgadStaNew.LONG = 0f;
                        asgadStaNew.NORTHING = "";
                        asgadStaNew.p = si.Prov;
                        asgadStaNew.PROV = si.Prov;
                        asgadStaNew.SECTOR = si.Sector;
                        asgadStaNew.SORT_ORDER = orderNumb.ToString();
                        asgadStaNew.STAT_NAME = "";
                        asgadStaNew.STAT_NBR = si.Stn_Nbr;
                        asgadStaNew.STAT_TYPE = "";
                        asgadStaNew.SUBSECTOR = si.Subsector;
                        asgadStaNew.UPDATED = null;
                        asgadStaNew.ZONE = null;

                        dbDT.ASGADStations.Add(asgadStaNew);
                        dbDT.SaveChanges();

                    }
                }
                lblStatus.Text = "Done fixing missing stations";
            }
        }
        private void LoadASGADMPNTABLE(string Prov)
        {
            List<ASGADMPNTable> asgadMPNTableList = new List<ASGADMPNTable>();

            string dirName = @"C:\ASGAD Data\" + Prov + @"\";

            string connectionString = @"Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + dirName + "; Extended Properties=DBASE III;";
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [MPNTABLE.dbf$]");

            OleDbDataReader reader;
            cmd.Connection = conn;
            reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            FieldNameList = new List<string>() { "TUBE5_10", "TUBE5_1", "TUBE5_01", "MPN" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                //richTextBox1.AppendText(reader.GetName(i) + " " + reader.GetValue(i).GetType() + "\r\n");
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            richTextBox1.AppendText("\r\n");

            reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                string TUBE5_10 = "";
                string TUBE5_1 = "";
                string TUBE5_01 = "";
                string MPN = "";

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    TUBE5_10 = "";
                }
                else
                {
                    TUBE5_10 = reader.GetValue(0).ToString().Trim();
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    TUBE5_1 = "";
                }
                else
                {
                    TUBE5_1 = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    TUBE5_01 = "";
                }
                else
                {
                    TUBE5_01 = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    MPN = "";
                }
                else
                {
                    MPN = reader.GetValue(3).ToString().Trim();
                }

                ASGADMPNTable asgadMPN = new ASGADMPNTable();
                asgadMPN.p = Prov;
                asgadMPN.TUBE5_10 = TUBE5_10;
                asgadMPN.TUBE5_1 = TUBE5_1;
                asgadMPN.TUBE5_01 = TUBE5_01;
                asgadMPN.MPN = MPN;

                asgadMPNTableList.Add(asgadMPN);
            }

            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (ASGADMPNTable ampn in asgadMPNTableList)
                {
                    dbDT.ASGADMPNTables.Add(ampn);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        private void LoadASGADPrecData(string Prov)
        {
            List<ASGADPrecData> asgadPrecDataList = new List<ASGADPrecData>();

            string dirName = @"C:\CSSP\ASGAD Data\" + Prov + @"\";

            string connectionString = @"Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + dirName + "; Extended Properties=DBASE III;";
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [precdata.dbf$]");

            OleDbDataReader reader;
            cmd.Connection = conn;
            reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            FieldNameList = new List<string>() { "PRECIPID", "SAMP_DATE", "PRECIP" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                //richTextBox1.AppendText(reader.GetName(i) + " " + reader.GetValue(i).GetType() + "\r\n");
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            richTextBox1.AppendText("\r\n");

            reader = cmd.ExecuteReader();

            int count = 0;
            while (reader.Read())
            {
                count += 1;
                lblTotal.Text = count.ToString();
                lblTotal.Refresh();
                Application.DoEvents();

                string PRECIPID = "";
                DateTime SAMP_DATE = new DateTime(2100, 1, 1);
                float? PRECIP = -1f;

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    PRECIPID = "";
                }
                else
                {
                    PRECIPID = reader.GetValue(0).ToString().Trim();
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    SAMP_DATE = new DateTime(2100, 1, 1);
                }
                else
                {
                    SAMP_DATE = (DateTime)(reader.GetValue(1));
                }
                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    PRECIP = null;
                }
                else
                {
                    PRECIP = (float)(double)(reader.GetValue(2));
                }

                ASGADPrecData asgadPrecData = new ASGADPrecData();
                asgadPrecData.p = Prov;
                asgadPrecData.PRECIPID = PRECIPID;
                if (Prov == "NS")
                {
                    if (SAMP_DATE == new DateTime(201, 9, 13) && PRECIPID == "8203400")
                    {
                        SAMP_DATE = new DateTime(2013, 9, 13);
                    }
                    if (SAMP_DATE == new DateTime(201, 9, 30) && PRECIPID == "8203400")
                    {
                        SAMP_DATE = new DateTime(2013, 9, 30);
                    }
                }
                asgadPrecData.SAMP_DATE = SAMP_DATE;
                asgadPrecData.PRECIP = PRECIP;

                asgadPrecDataList.Add(asgadPrecData);
            }

            int asgadPrecDataListCount = asgadPrecDataList.Count();
            int take = 200;
            int skip = 0;

            lblTotal.Text = asgadPrecDataList.Count.ToString();
            lblCountInDB.Text = "0";

            while (skip < asgadPrecDataListCount)
            {
                using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
                {
                    dbDT.ASGADPrecDatas.AddRange(asgadPrecDataList.Skip(skip).Take(take));
                    Application.DoEvents();
                    try
                    {
                        dbDT.SaveChanges();
                    }
                    catch (Exception)
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (ASGADPrecData apd in asgadPrecDataList.Skip(skip).Take(take))
                        {
                            sb.AppendLine(apd.PRECIP + "\t" + apd.PRECIPID + "\t" + apd.SAMP_DATE);
                        }
                        richTextBox1.Text = sb.ToString();
                        return;
                    }
                }

                lblCountInDB.Text = skip.ToString();
                lblCountInDB.Refresh();
                Application.DoEvents();

                skip += take;
            }
        }
        private void LoadASGADPrecStn(string Prov)
        {
            List<ASGADPrecStn> asgadPrecStnList = new List<ASGADPrecStn>();

            string dirName = @"C:\ASGAD Data\" + Prov + @"\";

            string connectionString = @"Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + dirName + "; Extended Properties=DBASE III;";
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [precstn.dbf$]");

            OleDbDataReader reader;
            cmd.Connection = conn;
            reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            FieldNameList = new List<string>() { "NAME", "PROV", "ID", "LAT", "LONG", "LAT_DEG", "LONG_DEG", "ELEV" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                //richTextBox1.AppendText(reader.GetName(i) + " " + reader.GetValue(i).GetType() + "\r\n");
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            richTextBox1.AppendText("\r\n");

            reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                string NAME = "";
                string PROV = "";
                string ID = "";
                string LAT = "";
                string LONG = "";
                float? LAT_DEG = -1f;
                float? LONG_DEG = -1f;
                float? ELEV = -1f;

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    NAME = "";
                }
                else
                {
                    NAME = reader.GetValue(0).ToString().Trim();
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    PROV = "";
                }
                else
                {
                    PROV = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    ID = "";
                }
                else
                {
                    ID = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    LONG = "";
                }
                else
                {
                    LONG = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    LAT = "";
                }
                else
                {
                    LAT = reader.GetValue(4).ToString().Trim();
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    LAT_DEG = null;
                }
                else
                {
                    LAT_DEG = (float)(double)(reader.GetValue(5));
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    LONG_DEG = null;
                }
                else
                {
                    LONG_DEG = (float)(double)(reader.GetValue(6));
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    ELEV = null;
                }
                else
                {
                    ELEV = (float)(double)(reader.GetValue(7));
                }

                ASGADPrecStn asgadPrecStn = new ASGADPrecStn();
                asgadPrecStn.NAME = NAME;
                asgadPrecStn.PROV = PROV;
                asgadPrecStn.ID = ID;
                asgadPrecStn.LAT = LAT;
                asgadPrecStn.LONG = LONG;
                if (LAT_DEG == null)
                {
                    asgadPrecStn.LAT_DEG = null;
                }
                else
                {
                    asgadPrecStn.LAT_DEG = (float)LAT_DEG;
                }
                if (LONG_DEG == null)
                {
                    asgadPrecStn.LONG_DEG = null;
                }
                else
                {
                    asgadPrecStn.LONG_DEG = (float)LONG_DEG;
                }
                if (ELEV == null)
                {
                    asgadPrecStn.ELEV = null;
                }
                else
                {
                    asgadPrecStn.ELEV = (float)ELEV;
                }

                asgadPrecStnList.Add(asgadPrecStn);
            }


            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (ASGADPrecStn aps in asgadPrecStnList)
                {
                    dbDT.ASGADPrecStns.Add(aps);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        private void LoadASGADRun(string Prov)
        {
            List<ASGADRun> asgadRunList = new List<ASGADRun>();

            string dirName = @"C:\CSSP\ASGAD Data\" + Prov + @"\";

            string connectionString = @"Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + dirName + "; Extended Properties=DBASE III;";
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [run.dbf$]");

            OleDbDataReader reader;
            cmd.Connection = conn;
            reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            FieldNameList = new List<string>() { "PROV", "AREA", "SECTOR", "SUBSECTOR", "SAMP_DATE", "RUN_NUMBER", "START_TIME", "FIN_TIME",
                "TIDE_CODE", "TIDEF_CODE", "AUTOCALC", "PPT24", "PPT48", "PPT72", "NOTE", "UPDATED" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                //richTextBox1.AppendText(reader.GetName(i) + " " + reader.GetValue(i).GetType() + "\r\n");
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            richTextBox1.AppendText("\r\n");

            reader = cmd.ExecuteReader();

            int count = 0;
            while (reader.Read())
            {
                count += 1;
                lblTotal.Text = count.ToString();
                lblTotal.Refresh();
                Application.DoEvents();

                string PROV = "";
                string AREA = "";
                string SECTOR = "";
                string SUBSECTOR = "";
                DateTime SAMP_DATE = new DateTime(2100, 1, 1);
                int? RUN_NUMBER = -1;
                string START_TIME = "";
                string FIN_TIME = "";
                string TIDE_CODE = "";
                string TIDEF_CODE = "";
                string AUTOCALC = "";
                float? PPT24 = -1f;
                float? PPT48 = -1f;
                float? PPT72 = -1f;
                string NOTE = "";
                DateTime UPDATED = new DateTime(2100, 1, 1);

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    PROV = "";
                }
                else
                {
                    PROV = reader.GetValue(0).ToString().Trim();
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    AREA = "";
                }
                else
                {
                    AREA = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    SECTOR = "";
                }
                else
                {
                    SECTOR = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    SUBSECTOR = "";
                }
                else
                {
                    SUBSECTOR = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    SAMP_DATE = new DateTime(2100, 1, 1);
                }
                else
                {
                    SAMP_DATE = (DateTime)(reader.GetValue(4));
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    RUN_NUMBER = null;
                }
                else
                {
                    RUN_NUMBER = int.Parse(reader.GetValue(5).ToString().Trim());
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    START_TIME = null;
                }
                else
                {
                    START_TIME = reader.GetValue(6).ToString().Trim();
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    FIN_TIME = null;
                }
                else
                {
                    FIN_TIME = reader.GetValue(7).ToString().Trim();
                }

                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    TIDE_CODE = null;
                }
                else
                {
                    TIDE_CODE = reader.GetValue(8).ToString().Trim();
                }

                if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                {
                    TIDEF_CODE = null;
                }
                else
                {
                    TIDEF_CODE = reader.GetValue(9).ToString().Trim();
                }

                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                {
                    AUTOCALC = null;
                }
                else
                {
                    AUTOCALC = reader.GetValue(10).ToString().Trim();
                }

                if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                {
                    PPT24 = null;
                }
                else
                {
                    PPT24 = (float)(double)(reader.GetValue(11));
                }

                if (reader.GetValue(12).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(12).ToString()))
                {
                    PPT48 = null;
                }
                else
                {
                    PPT48 = (float)(double)(reader.GetValue(12));
                }

                if (reader.GetValue(13).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(13).ToString()))
                {
                    PPT72 = null;
                }
                else
                {
                    PPT72 = (float)(double)(reader.GetValue(13));
                }

                if (reader.GetValue(14).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(14).ToString()))
                {
                    NOTE = null;
                }
                else
                {
                    NOTE = reader.GetValue(14).ToString().Trim();
                }

                if (reader.GetValue(15).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(15).ToString()))
                {
                    UPDATED = new DateTime(2100, 1, 1);
                }
                else
                {
                    UPDATED = (DateTime)(reader.GetValue(15));
                }

                ASGADRun asgadRun = new ASGADRun();
                asgadRun.p = Prov;
                asgadRun.PROV = PROV;
                asgadRun.AREA = AREA;
                asgadRun.SECTOR = SECTOR;
                asgadRun.SUBSECTOR = SUBSECTOR;
                asgadRun.SAMP_DATE = SAMP_DATE;
                asgadRun.RUN_NUMBER = RUN_NUMBER;
                asgadRun.START_TIME = START_TIME;
                asgadRun.FIN_TIME = FIN_TIME;
                asgadRun.TIDE_CODE = TIDE_CODE;
                asgadRun.TIDEF_CODE = TIDEF_CODE;
                asgadRun.AUTOCALC = AUTOCALC;
                asgadRun.PPT24 = PPT24;
                asgadRun.PPT48 = PPT48;
                asgadRun.PPT72 = PPT72;
                asgadRun.NOTE = NOTE;
                asgadRun.UPDATED = UPDATED;

                asgadRunList.Add(asgadRun);
            }

            int asgadRunListCount = asgadRunList.Count();
            int take = 200;
            int skip = 0;

            lblTotal.Text = asgadRunList.Count.ToString();
            lblCountInDB.Text = "0";

            while (skip < asgadRunListCount)
            {
                using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
                {
                    dbDT.ASGADRuns.AddRange(asgadRunList.Skip(skip).Take(take));
                    Application.DoEvents();
                    dbDT.SaveChanges();
                }

                lblCountInDB.Text = count.ToString();
                lblCountInDB.Refresh();
                Application.DoEvents();


                lblCountInDB.Text = skip.ToString();
                skip += take;
            }
        }
        private void LoadASGADRunMOU(string Prov)
        {
            List<ASGADRunMOU> asgadRunMOUList = new List<ASGADRunMOU>();

            string dirName = @"C:\CSSP Latest Code Old\DataTool\NBMOUDATA\" + Prov + @"\";

            string connectionString = @"Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + dirName + "; Extended Properties=DBASE III;";
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [run.dbf$]");

            OleDbDataReader reader;
            cmd.Connection = conn;
            reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            FieldNameList = new List<string>() { "PROV", "AREA", "SECTOR", "SUBSECTOR", "SAMP_DATE", "RUN_NUMBER", "START_TIME", "FIN_TIME",
                "TIDE_CODE", "TIDEF_CODE", "AUTOCALC", "PPT24", "PPT48", "PPT72", "NOTE", "UPDATED" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                //richTextBox1.AppendText(reader.GetName(i) + " " + reader.GetValue(i).GetType() + "\r\n");
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            richTextBox1.AppendText("\r\n");

            reader = cmd.ExecuteReader();

            int count = 0;
            while (reader.Read())
            {
                count += 1;
                lblTotal.Text = count.ToString();
                lblTotal.Refresh();
                Application.DoEvents();

                string PROV = "";
                string AREA = "";
                string SECTOR = "";
                string SUBSECTOR = "";
                DateTime SAMP_DATE = new DateTime(2100, 1, 1);
                int? RUN_NUMBER = -1;
                string START_TIME = "";
                string FIN_TIME = "";
                string TIDE_CODE = "";
                string TIDEF_CODE = "";
                string AUTOCALC = "";
                float? PPT24 = -1f;
                float? PPT48 = -1f;
                float? PPT72 = -1f;
                string NOTE = "";
                DateTime UPDATED = new DateTime(2100, 1, 1);

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    PROV = "";
                }
                else
                {
                    PROV = reader.GetValue(0).ToString().Trim();
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    AREA = "";
                }
                else
                {
                    AREA = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    SECTOR = "";
                }
                else
                {
                    SECTOR = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    SUBSECTOR = "";
                }
                else
                {
                    SUBSECTOR = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    SAMP_DATE = new DateTime(2100, 1, 1);
                }
                else
                {
                    SAMP_DATE = (DateTime)(reader.GetValue(4));
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    RUN_NUMBER = null;
                }
                else
                {
                    RUN_NUMBER = int.Parse(reader.GetValue(5).ToString().Trim());
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    START_TIME = null;
                }
                else
                {
                    START_TIME = reader.GetValue(6).ToString().Trim();
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    FIN_TIME = null;
                }
                else
                {
                    FIN_TIME = reader.GetValue(7).ToString().Trim();
                }

                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    TIDE_CODE = null;
                }
                else
                {
                    TIDE_CODE = reader.GetValue(8).ToString().Trim();
                }

                if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                {
                    TIDEF_CODE = null;
                }
                else
                {
                    TIDEF_CODE = reader.GetValue(9).ToString().Trim();
                }

                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                {
                    AUTOCALC = null;
                }
                else
                {
                    AUTOCALC = reader.GetValue(10).ToString().Trim();
                }

                if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                {
                    PPT24 = null;
                }
                else
                {
                    PPT24 = (float)(double)(reader.GetValue(11));
                }

                if (reader.GetValue(12).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(12).ToString()))
                {
                    PPT48 = null;
                }
                else
                {
                    PPT48 = (float)(double)(reader.GetValue(12));
                }

                if (reader.GetValue(13).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(13).ToString()))
                {
                    PPT72 = null;
                }
                else
                {
                    PPT72 = (float)(double)(reader.GetValue(13));
                }

                if (reader.GetValue(14).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(14).ToString()))
                {
                    NOTE = null;
                }
                else
                {
                    NOTE = reader.GetValue(14).ToString().Trim();
                }

                if (reader.GetValue(15).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(15).ToString()))
                {
                    UPDATED = new DateTime(2100, 1, 1);
                }
                else
                {
                    UPDATED = (DateTime)(reader.GetValue(15));
                }

                ASGADRunMOU asgadRunMOU = new ASGADRunMOU();
                asgadRunMOU.p = Prov;
                asgadRunMOU.PROV = PROV;
                asgadRunMOU.AREA = AREA;
                asgadRunMOU.SECTOR = SECTOR;
                asgadRunMOU.SUBSECTOR = SUBSECTOR;
                asgadRunMOU.SAMP_DATE = SAMP_DATE;
                asgadRunMOU.RUN_NUMBER = RUN_NUMBER;
                asgadRunMOU.START_TIME = START_TIME;
                asgadRunMOU.FIN_TIME = FIN_TIME;
                asgadRunMOU.TIDE_CODE = TIDE_CODE;
                asgadRunMOU.TIDEF_CODE = TIDEF_CODE;
                asgadRunMOU.AUTOCALC = AUTOCALC;
                asgadRunMOU.PPT24 = PPT24;
                asgadRunMOU.PPT48 = PPT48;
                asgadRunMOU.PPT72 = PPT72;
                asgadRunMOU.NOTE = NOTE;
                asgadRunMOU.UPDATED = UPDATED;

                asgadRunMOUList.Add(asgadRunMOU);
            }

            int asgadRunMOUListCount = asgadRunMOUList.Count();
            int take = 200;
            int skip = 0;

            lblTotal.Text = asgadRunMOUList.Count.ToString();
            lblCountInDB.Text = "0";

            while (skip < asgadRunMOUListCount)
            {
                using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
                {
                    dbDT.ASGADRunMOUs.AddRange(asgadRunMOUList.Skip(skip).Take(take));
                    Application.DoEvents();
                    dbDT.SaveChanges();
                }

                lblCountInDB.Text = count.ToString();
                lblCountInDB.Refresh();
                Application.DoEvents();


                lblCountInDB.Text = skip.ToString();
                skip += take;
            }
        }
        private void LoadASGADSampCode(string Prov)
        {
            List<ASGADSampCode> asgadSampCodeList = new List<ASGADSampCode>();

            string dirName = @"C:\ASGAD Data\" + Prov + @"\";

            string connectionString = @"Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + dirName + "; Extended Properties=DBASE III;";
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [sampcode.dbf$]");

            OleDbDataReader reader;
            cmd.Connection = conn;
            reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            FieldNameList = new List<string>() { "SAMPCODE", "OTH_CODE", "SAMPNAME", "UNITS", "LOW_LIMIT", "HIGH_LIMIT" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                //richTextBox1.AppendText(reader.GetName(i) + " " + reader.GetValue(i).GetType() + "\r\n");
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            richTextBox1.AppendText("\r\n");

            reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                string SAMPCODE = "";
                string OTH_CODE = "";
                string SAMPNAME = "";
                string UNITS = "";
                string LOW_LIMIT = "";
                string HIGH_LIMIT = "";

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    SAMPCODE = "";
                }
                else
                {
                    SAMPCODE = reader.GetValue(0).ToString().Trim();
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    OTH_CODE = "";
                }
                else
                {
                    OTH_CODE = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    SAMPNAME = "";
                }
                else
                {
                    SAMPNAME = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    UNITS = "";
                }
                else
                {
                    UNITS = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    LOW_LIMIT = null;
                }
                else
                {
                    LOW_LIMIT = reader.GetValue(4).ToString().Trim();
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    HIGH_LIMIT = null;
                }
                else
                {
                    HIGH_LIMIT = reader.GetValue(5).ToString().Trim();
                }

                ASGADSampCode asgadSampCode = new ASGADSampCode();
                asgadSampCode.p = Prov;
                asgadSampCode.SAMPCODE = SAMPCODE;
                asgadSampCode.OTH_CODE = OTH_CODE;
                asgadSampCode.SAMPNAME = SAMPNAME;
                asgadSampCode.UNITS = UNITS;
                asgadSampCode.LOW_LIMIT = LOW_LIMIT;
                asgadSampCode.HIGH_LIMIT = HIGH_LIMIT;

                asgadSampCodeList.Add(asgadSampCode);
            }

            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                foreach (ASGADSampCode asc in asgadSampCodeList)
                {
                    dbDT.ASGADSampCodes.Add(asc);
                    Application.DoEvents();
                }
                dbDT.SaveChanges();
            }
        }
        private void LoadASGADSample(string Prov)
        {
            List<ASGADSample> asgadSampleList = new List<ASGADSample>();

            string dirName = @"C:\CSSP\ASGAD Data\" + Prov + @"\";

            string connectionString = @"Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + dirName + "; Extended Properties=DBASE III;";
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [sample.dbf$]");

            OleDbDataReader reader;
            cmd.Connection = conn;
            reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            FieldNameList = new List<string>() { "PROV", "AREA", "SECTOR", "SUBSECTOR", "SAMP_DATE", "RUN_NUMBER", "STAT_NBR", "SAMP_TIME", "FEC_COL", "SALINITY", "WATERTEMP", "UPDATED" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                //richTextBox1.AppendText(reader.GetName(i) + " " + reader.GetValue(i).GetType() + "\r\n");
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            richTextBox1.AppendText("\r\n");

            reader = cmd.ExecuteReader();

            int count = 0;
            while (reader.Read())
            {
                count += 1;

                lblTotal.Text = count.ToString();
                lblTotal.Refresh();
                Application.DoEvents();

                string PROV = "";
                string AREA = "";
                string SECTOR = "";
                string SUBSECTOR = "";
                DateTime SAMP_DATE = new DateTime(2100, 1, 1);
                int? RUN_NUMBER = -1;
                string STAT_NBR = "";
                string SAMP_TIME = "";
                float? FEC_COL = -1f;
                float? SALINITY = -1f;
                float? WATERTEMP = -1f;
                DateTime UPDATED = new DateTime(2100, 1, 1);

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    PROV = "";
                }
                else
                {
                    PROV = reader.GetValue(0).ToString().Trim();
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    AREA = "";
                }
                else
                {
                    AREA = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    SECTOR = "";
                }
                else
                {
                    SECTOR = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    SUBSECTOR = "";
                }
                else
                {
                    SUBSECTOR = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    SAMP_DATE = new DateTime(2100, 1, 1);
                }
                else
                {
                    SAMP_DATE = (DateTime)(reader.GetValue(4));
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    RUN_NUMBER = null;
                }
                else
                {
                    RUN_NUMBER = int.Parse(reader.GetValue(5).ToString().Trim());
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    STAT_NBR = null;
                }
                else
                {
                    STAT_NBR = reader.GetValue(6).ToString().Trim();
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    SAMP_TIME = null;
                }
                else
                {
                    SAMP_TIME = reader.GetValue(7).ToString().Trim();
                }

                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    FEC_COL = null;
                }
                else
                {
                    FEC_COL = (float)(double)(reader.GetValue(8));
                }

                if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                {
                    SALINITY = null;
                }
                else
                {
                    SALINITY = (float)(double)(reader.GetValue(9));
                }

                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                {
                    WATERTEMP = null;
                }
                else
                {
                    WATERTEMP = (float)(double)(reader.GetValue(10));
                }

                if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                {
                    UPDATED = new DateTime(2100, 1, 1);
                }
                else
                {
                    UPDATED = (DateTime)(reader.GetValue(11));
                }

                ASGADSample asgadSample = new ASGADSample();
                asgadSample.p = Prov;
                asgadSample.PROV = PROV;
                asgadSample.AREA = AREA;
                asgadSample.SECTOR = SECTOR;
                asgadSample.SUBSECTOR = SUBSECTOR;
                asgadSample.SAMP_DATE = SAMP_DATE;
                asgadSample.RUN_NUMBER = RUN_NUMBER;
                asgadSample.STAT_NBR = STAT_NBR;
                asgadSample.SAMP_TIME = SAMP_TIME;
                asgadSample.FEC_COL = FEC_COL;
                asgadSample.SALINITY = SALINITY;
                asgadSample.WATERTEMP = WATERTEMP;
                asgadSample.UPDATED = UPDATED;

                asgadSampleList.Add(asgadSample);
            }
            int asgadSampleListCount = asgadSampleList.Count();
            int take = 200;
            int skip = 0;

            lblTotal.Text = asgadSampleList.Count.ToString();
            lblCountInDB.Text = "0";

            while (skip < asgadSampleListCount)
            {
                using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
                {
                    dbDT.ASGADSamples.AddRange(asgadSampleList.Skip(skip).Take(take));
                    Application.DoEvents();
                    dbDT.SaveChanges();
                }

                lblCountInDB.Text = skip.ToString();
                lblCountInDB.Refresh();
                Application.DoEvents();

                skip += take;
            }

        }
        private void LoadASGADSampleMOU(string Prov)
        {
            List<ASGADSampleMOU> asgadSampleMOUList = new List<ASGADSampleMOU>();

            string dirName = @"C:\CSSP Latest Code Old\DataTool\NBMOUDATA\" + Prov + @"\";

            string connectionString = @"Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + dirName + "; Extended Properties=DBASE III;";
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [sample.dbf$]");

            OleDbDataReader reader;
            cmd.Connection = conn;
            reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            FieldNameList = new List<string>() { "PROV", "AREA", "SECTOR", "SUBSECTOR", "SAMP_DATE", "RUN_NUMBER", "STAT_NBR", "SAMP_TIME", "FEC_COL", "SALINITY", "WATERTEMP", "UPDATED" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                //richTextBox1.AppendText(reader.GetName(i) + " " + reader.GetValue(i).GetType() + "\r\n");
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            richTextBox1.AppendText("\r\n");

            reader = cmd.ExecuteReader();

            int count = 0;
            while (reader.Read())
            {
                count += 1;

                lblTotal.Text = count.ToString();
                lblTotal.Refresh();
                Application.DoEvents();

                string PROV = "";
                string AREA = "";
                string SECTOR = "";
                string SUBSECTOR = "";
                DateTime SAMP_DATE = new DateTime(2100, 1, 1);
                int? RUN_NUMBER = -1;
                string STAT_NBR = "";
                string SAMP_TIME = "";
                float? FEC_COL = -1f;
                float? SALINITY = -1f;
                float? WATERTEMP = -1f;
                DateTime UPDATED = new DateTime(2100, 1, 1);

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    PROV = "";
                }
                else
                {
                    PROV = reader.GetValue(0).ToString().Trim();
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    AREA = "";
                }
                else
                {
                    AREA = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    SECTOR = "";
                }
                else
                {
                    SECTOR = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    SUBSECTOR = "";
                }
                else
                {
                    SUBSECTOR = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    SAMP_DATE = new DateTime(2100, 1, 1);
                }
                else
                {
                    SAMP_DATE = (DateTime)(reader.GetValue(4));
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    RUN_NUMBER = null;
                }
                else
                {
                    RUN_NUMBER = int.Parse(reader.GetValue(5).ToString().Trim());
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    STAT_NBR = null;
                }
                else
                {
                    STAT_NBR = reader.GetValue(6).ToString().Trim();
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    SAMP_TIME = null;
                }
                else
                {
                    SAMP_TIME = reader.GetValue(7).ToString().Trim();
                }

                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    FEC_COL = null;
                }
                else
                {
                    FEC_COL = (float)(double)(reader.GetValue(8));
                }

                if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                {
                    SALINITY = null;
                }
                else
                {
                    SALINITY = (float)(double)(reader.GetValue(9));
                }

                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                {
                    WATERTEMP = null;
                }
                else
                {
                    WATERTEMP = (float)(double)(reader.GetValue(10));
                }

                if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                {
                    UPDATED = new DateTime(2100, 1, 1);
                }
                else
                {
                    UPDATED = (DateTime)(reader.GetValue(11));
                }

                ASGADSampleMOU asgadSampleMOU = new ASGADSampleMOU();
                asgadSampleMOU.p = Prov;
                asgadSampleMOU.PROV = PROV;
                asgadSampleMOU.AREA = AREA;
                asgadSampleMOU.SECTOR = SECTOR;
                asgadSampleMOU.SUBSECTOR = SUBSECTOR;
                asgadSampleMOU.SAMP_DATE = SAMP_DATE;
                asgadSampleMOU.RUN_NUMBER = RUN_NUMBER;
                asgadSampleMOU.STAT_NBR = STAT_NBR;
                asgadSampleMOU.SAMP_TIME = SAMP_TIME;
                asgadSampleMOU.FEC_COL = FEC_COL;
                asgadSampleMOU.SALINITY = SALINITY;
                asgadSampleMOU.WATERTEMP = WATERTEMP;
                asgadSampleMOU.UPDATED = UPDATED;

                asgadSampleMOUList.Add(asgadSampleMOU);
            }
            int asgadSampleListCount = asgadSampleMOUList.Count();
            int take = 200;
            int skip = 0;

            lblTotal.Text = asgadSampleMOUList.Count.ToString();
            lblCountInDB.Text = "0";

            while (skip < asgadSampleListCount)
            {
                using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
                {
                    dbDT.ASGADSampleMOUs.AddRange(asgadSampleMOUList.Skip(skip).Take(take));
                    Application.DoEvents();
                    dbDT.SaveChanges();
                }

                lblCountInDB.Text = skip.ToString();
                lblCountInDB.Refresh();
                Application.DoEvents();

                skip += take;
            }

        }
        private void LoadASGADStation(string Prov)
        {
            List<ASGADStation> asgadStationList = new List<ASGADStation>();

            string dirName = @"C:\CSSP\ASGAD Data\" + Prov + @"\";

            string connectionString = @"Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + dirName + "; Extended Properties=DBASE III;";
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [station.dbf$]");

            OleDbDataReader reader;
            cmd.Connection = conn;
            reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            FieldNameList = new List<string>() { "PROV", "AREA", "SECTOR", "SUBSECTOR", "STAT_NBR", "STAT_NAME", "STAT_TYPE", "SORT_ORDER", "ZONE",
                "EASTING", "NORTHING", "LAT", "LONG", "ACTIVE", "UPDATED" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                //richTextBox1.AppendText(reader.GetName(i) + " " + reader.GetValue(i).GetType() + "\r\n");
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();

            reader = cmd.ExecuteReader();

            int count = 0;
            while (reader.Read())
            {
                count += 1;
                lblTotal.Text = count.ToString();
                lblTotal.Refresh();
                Application.DoEvents();

                string PROV = "";
                string AREA = "";
                string SECTOR = "";
                string SUBSECTOR = "";
                string STAT_NBR = "";
                string STAT_NAME = "";
                string STAT_TYPE = "";
                string SORT_ORDER = "";
                string ZONE = "";
                string EASTING = "";
                string NORTHING = "";
                float? LAT = -1;
                float? LONG = -1f;
                string ACTIVE = "";
                DateTime UPDATED = new DateTime(2100, 1, 1);

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    PROV = "";
                }
                else
                {
                    PROV = reader.GetValue(0).ToString().Trim();
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    AREA = "";
                }
                else
                {
                    AREA = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    SECTOR = "";
                }
                else
                {
                    SECTOR = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    SUBSECTOR = "";
                }
                else
                {
                    SUBSECTOR = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    STAT_NBR = "";
                }
                else
                {
                    STAT_NBR = reader.GetValue(4).ToString().Trim();
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    STAT_NAME = "";
                }
                else
                {
                    STAT_NAME = reader.GetValue(5).ToString().Trim();
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    STAT_TYPE = "";
                }
                else
                {
                    STAT_TYPE = reader.GetValue(6).ToString().Trim();
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    SORT_ORDER = null;
                }
                else
                {
                    SORT_ORDER = reader.GetValue(7).ToString().Trim();
                }

                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    ZONE = null;
                }
                else
                {
                    ZONE = reader.GetValue(8).ToString().Trim();
                }

                if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                {
                    EASTING = null;
                }
                else
                {
                    EASTING = reader.GetValue(9).ToString().Trim();
                }

                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                {
                    NORTHING = null;
                }
                else
                {
                    NORTHING = reader.GetValue(10).ToString().Trim();
                }

                if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                {
                    LAT = null;
                }
                else
                {
                    LAT = (float)(double)(reader.GetValue(11));
                }

                if (reader.GetValue(12).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(12).ToString()))
                {
                    LONG = null;
                }
                else
                {
                    LONG = (float)(double)(reader.GetValue(12));
                }

                if (reader.GetValue(13).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(13).ToString()))
                {
                    ACTIVE = "";
                }
                else
                {
                    ACTIVE = reader.GetValue(13).ToString().Trim();
                }

                if (reader.GetValue(14).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(14).ToString()))
                {
                    UPDATED = new DateTime(2100, 1, 1);
                }
                else
                {
                    UPDATED = (DateTime)(reader.GetValue(14));
                }

                ASGADStation asgadStation = new ASGADStation();
                asgadStation.p = Prov;
                asgadStation.PROV = PROV;
                asgadStation.AREA = AREA;
                asgadStation.SECTOR = SECTOR;
                asgadStation.SUBSECTOR = SUBSECTOR;
                asgadStation.STAT_NBR = STAT_NBR;
                asgadStation.STAT_NAME = STAT_NAME;
                asgadStation.STAT_TYPE = STAT_TYPE;
                asgadStation.SORT_ORDER = SORT_ORDER;
                asgadStation.ZONE = ZONE;
                asgadStation.EASTING = EASTING;
                asgadStation.NORTHING = NORTHING;
                asgadStation.LAT = LAT;
                asgadStation.LONG = LONG;
                asgadStation.ACTIVE = ACTIVE;
                asgadStation.UPDATED = UPDATED;

                asgadStationList.Add(asgadStation);
            }

            int asgadStationListCount = asgadStationList.Count();
            int take = 200;
            int skip = 0;

            lblTotal.Text = asgadStationList.Count.ToString();
            lblCountInDB.Text = "0";

            while (skip < asgadStationListCount)
            {
                using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
                {
                    dbDT.ASGADStations.AddRange(asgadStationList.Skip(skip).Take(take));
                    Application.DoEvents();
                    dbDT.SaveChanges();
                }

                lblCountInDB.Text = skip.ToString();
                lblCountInDB.Refresh();
                Application.DoEvents();

                skip += take;
            }

        }
        private void LoadASGADSubsect(string Prov)
        {
            List<ASGADSubsect> asgadSubsectList = new List<ASGADSubsect>();

            string dirName = @"C:\ASGAD Data\" + Prov + @"\";

            string connectionString = @"Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + dirName + "; Extended Properties=DBASE III;";
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [subsect.dbf$] where [PROV] = '" + Prov + "'");

            OleDbDataReader reader;
            cmd.Connection = conn;
            reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            FieldNameList = new List<string>() { "PROV", "AREA", "SECTOR", "SUBSECTOR", "DFO_CODE", "NAME", "DESCRIPT", "MAP",
                "PRECIPID", "PRECIPID2", "PRECIPID3", "UPDATED" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                //richTextBox1.AppendText(reader.GetName(i) + " " + reader.GetValue(i).GetType() + "\r\n");
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            richTextBox1.AppendText("\r\n");

            reader = cmd.ExecuteReader();

            int count = 0;
            while (reader.Read())
            {
                count += 1;
                lblTotal.Text = count.ToString();
                lblTotal.Refresh();
                Application.DoEvents();

                string PROV = "";
                string AREA = "";
                string SECTOR = "";
                string SUBSECTOR = "";
                string DFO_CODE = "";
                string NAME = "";
                string DESCRIPT = "";
                string MAP = "";
                string PRECIPID = "";
                string PRECIPID2 = "";
                string PRECIPID3 = "";
                DateTime UPDATED = new DateTime(2050, 1, 1);

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    PROV = "";
                }
                else
                {
                    PROV = reader.GetValue(0).ToString().Trim();
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    AREA = "";
                }
                else
                {
                    AREA = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    SECTOR = "";
                }
                else
                {
                    SECTOR = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    SUBSECTOR = "";
                }
                else
                {
                    SUBSECTOR = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    DFO_CODE = null;
                }
                else
                {
                    DFO_CODE = reader.GetValue(4).ToString().Trim();
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    NAME = null;
                }
                else
                {
                    NAME = reader.GetValue(5).ToString().Trim();
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    DESCRIPT = null;
                }
                else
                {
                    DESCRIPT = reader.GetValue(6).ToString().Trim();
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    MAP = null;
                }
                else
                {
                    MAP = reader.GetValue(7).ToString().Trim();
                }


                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    PRECIPID = null;
                }
                else
                {
                    PRECIPID = reader.GetValue(8).ToString().Trim();
                }

                if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                {
                    PRECIPID2 = null;
                }
                else
                {
                    PRECIPID2 = reader.GetValue(9).ToString().Trim();
                }

                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                {
                    PRECIPID3 = null;
                }
                else
                {
                    PRECIPID3 = reader.GetValue(10).ToString().Trim();
                }

                if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                {
                    UPDATED = new DateTime(2050, 1, 1);
                }
                else
                {
                    UPDATED = (DateTime)(reader.GetValue(11));
                }

                ASGADSubsect asgadSubsect = new ASGADSubsect();
                asgadSubsect.p = Prov;
                asgadSubsect.PROV = PROV;
                asgadSubsect.AREA = AREA;
                asgadSubsect.SECTOR = SECTOR;
                asgadSubsect.SUBSECTOR = SUBSECTOR;
                asgadSubsect.DFO_CODE = DFO_CODE;
                asgadSubsect.NAME = NAME;
                asgadSubsect.DESCRIPT = DESCRIPT;
                asgadSubsect.MAP = MAP;
                asgadSubsect.PRECIPID = PRECIPID;
                asgadSubsect.PRECIPID2 = PRECIPID2;
                asgadSubsect.PRECIPID3 = PRECIPID3;
                asgadSubsect.UPDATED = UPDATED;

                asgadSubsectList.Add(asgadSubsect);
            }

            int asgadSubsectListCount = asgadSubsectList.Count();
            int take = 100;
            int skip = 0;

            lblTotal.Text = asgadSubsectList.Count.ToString();
            lblCountInDB.Text = "0";

            while (skip < asgadSubsectListCount)
            {
                using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
                {
                    foreach (ASGADSubsect ass in asgadSubsectList.Skip(skip).Take(take))
                    {
                        dbDT.ASGADSubsects.Add(ass);
                        Application.DoEvents();
                    }
                    dbDT.SaveChanges();
                }

                lblCountInDB.Text = skip.ToString();
                lblCountInDB.Refresh();
                Application.DoEvents();

                skip += take;
            }

        }
        private void LoadASGADTide(string Prov)
        {

            List<ASGADTide> asgadTideList = new List<ASGADTide>();

            string dirName = @"C:\ASGAD Data\" + Prov + @"\";

            string connectionString = @"Provider=Microsoft.Jet.OleDb.4.0; Data Source=" + dirName + "; Extended Properties=DBASE III;";
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [tide.dbf$]");

            OleDbDataReader reader;
            cmd.Connection = conn;
            reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            FieldNameList = new List<string>() { "TIDE_CODE", "TIDE_NAME" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                //richTextBox1.AppendText(reader.GetName(i) + " " + reader.GetValue(i).GetType() + "\r\n");
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            richTextBox1.AppendText("\r\n");

            reader = cmd.ExecuteReader();

            int count = 0;
            while (reader.Read())
            {
                count += 1;
                lblTotal.Text = count.ToString();
                lblTotal.Refresh();
                Application.DoEvents();

                string TIDE_CODE = "";
                string TIDE_NAME = "";

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    TIDE_CODE = "";
                }
                else
                {
                    TIDE_CODE = reader.GetValue(0).ToString().Trim();
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    TIDE_NAME = "";
                }
                else
                {
                    TIDE_NAME = reader.GetValue(1).ToString().Trim();
                }


                ASGADTide asgadTide = new ASGADTide();
                asgadTide.p = Prov;
                asgadTide.TIDE_CODE = TIDE_CODE;
                asgadTide.TIDE_NAME = TIDE_NAME;

                asgadTideList.Add(asgadTide);
            }

            int asgadTideListCount = asgadTideList.Count();
            int take = 100;
            int skip = 0;

            lblTotal.Text = asgadTideList.Count.ToString();
            lblCountInDB.Text = "0";

            while (skip < asgadTideListCount)
            {
                using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
                {
                    foreach (ASGADTide ast in asgadTideList.Skip(skip).Take(take))
                    {
                        dbDT.ASGADTides.Add(ast);
                        Application.DoEvents();
                    }
                    dbDT.SaveChanges();
                }

                lblCountInDB.Text = skip.ToString();
                lblCountInDB.Refresh();
                Application.DoEvents();

                skip += take;
            }
        }
        private void LoadBCLand_Base_Sample_Readings()
        {
            List<BCLandSample> bcLandSampleList = new List<BCLandSample>();

            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OleDb.12.0;Data Source=C:\CSSP\BC_Data\CSSP_Pacific_WQ_Data_20180219.accdb");
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Land Based Sample Reading]", conn);
            OleDbDataReader reader = cmd.ExecuteReader();

            List<string> FieldNameList = new List<string>();

            FieldNameList = new List<string>() { "ID", "SR_SURVEY", "SR_STATION_CODE", "SR_STATION_TYPE",
                "SR_ANALYSIS_TYPE", "SR_READING_DATE", "SR_READING_TIME", "SR_SAMPLE_TYPE", "SR_FECAL_COLIFORM_IND",
                "SR_FECAL_COLIFORM", "SR_ENTEROCOCCI_IND", "SR_ENTEROCOCCI", "SR_FLOW", "SR_SAMPLE_AGENCY",
                "SR_OLD_KEY" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            reader = cmd.ExecuteReader();

            int Count = 0;
            while (reader.Read())
            {
                Count += 1;
                Application.DoEvents();

                int? ID = -1;
                int? SR_SURVEY = -1;
                string SR_STATION_CODE = "";
                string SR_STATION_TYPE = "";
                string SR_ANALYSIS_TYPE = "";
                DateTime SR_READING_DATE = new DateTime(2050, 1, 1);
                string SR_READING_TIME = "";
                string SR_SAMPLE_TYPE = "";
                string SR_FECAL_COLIFORM_IND = "";
                int? SR_FECAL_COLIFORM = -1;
                string SR_ENTEROCOCCI_IND = "";
                int? SR_ENTEROCOCCI = -1;
                string SR_FLOW = "";
                int? SR_SAMPLE_AGENCY = -1;
                string SR_OLD_KEY = "";
                float? SR_SAMPLE_DEPTH = -1;
                float? SR_SALINITY = -1;
                float? SR_TEMPERATURE = -1;
                string SR_TIDE_CODE = "";
                string SR_OBS = "";

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    ID = null;
                }
                else
                {
                    ID = (int)(reader.GetValue(0));
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    SR_SURVEY = null;
                }
                else
                {
                    SR_SURVEY = (int)(double)(reader.GetValue(1));
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    SR_STATION_CODE = null;
                }
                else
                {
                    SR_STATION_CODE = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    SR_STATION_TYPE = null;
                }
                else
                {
                    SR_STATION_TYPE = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    SR_ANALYSIS_TYPE = null;
                }
                else
                {
                    SR_ANALYSIS_TYPE = reader.GetValue(4).ToString().Trim();
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    SR_READING_DATE = new DateTime(2050, 1, 1);
                }
                else
                {
                    SR_READING_DATE = (DateTime)(reader.GetValue(5));
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    SR_READING_TIME = null;
                }
                else
                {
                    SR_READING_TIME = reader.GetValue(6).ToString().Trim();
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    SR_SAMPLE_TYPE = null;
                }
                else
                {
                    SR_SAMPLE_TYPE = reader.GetValue(7).ToString().Trim();
                }

                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    SR_FECAL_COLIFORM_IND = null;
                }
                else
                {
                    SR_FECAL_COLIFORM_IND = reader.GetValue(8).ToString().Trim();
                }

                if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                {
                    SR_FECAL_COLIFORM = null;
                }
                else
                {
                    SR_FECAL_COLIFORM = (int)(double)(reader.GetValue(9));
                }

                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                {
                    SR_ENTEROCOCCI_IND = null;
                }
                else
                {
                    SR_ENTEROCOCCI_IND = reader.GetValue(10).ToString().Trim();
                }

                if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                {
                    SR_ENTEROCOCCI = null;
                }
                else
                {
                    SR_ENTEROCOCCI = (int)(reader.GetValue(11));
                }

                if (reader.GetValue(12).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(12).ToString()))
                {
                    SR_FLOW = null;
                }
                else
                {
                    SR_FLOW = reader.GetValue(12).ToString().Trim();
                }

                if (reader.GetValue(13).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(13).ToString()))
                {
                    SR_SAMPLE_AGENCY = null;
                }
                else
                {
                    SR_SAMPLE_AGENCY = (int)(double)(reader.GetValue(13));
                }

                if (reader.GetValue(14).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(14).ToString()))
                {
                    SR_OLD_KEY = null;
                }
                else
                {
                    SR_OLD_KEY = reader.GetValue(14).ToString().Trim();
                }

               

                BCLandSample bcLandSampleNew = new BCLandSample();
                bcLandSampleNew.ID = ID;
                bcLandSampleNew.SR_SURVEY = SR_SURVEY;
                bcLandSampleNew.SR_STATION_CODE = SR_STATION_CODE;
                bcLandSampleNew.SR_STATION_TYPE = SR_STATION_TYPE;
                bcLandSampleNew.SR_ANALYSIS_TYPE = SR_ANALYSIS_TYPE;
                bcLandSampleNew.SR_READING_DATE = SR_READING_DATE;
                bcLandSampleNew.SR_READING_TIME = SR_READING_TIME;
                bcLandSampleNew.SR_SAMPLE_TYPE = SR_SAMPLE_TYPE;
                bcLandSampleNew.SR_FECAL_COLIFORM_IND = SR_FECAL_COLIFORM_IND;
                bcLandSampleNew.SR_FECAL_COLIFORM = SR_FECAL_COLIFORM;
                bcLandSampleNew.SR_ENTEROCOCCI_IND = SR_ENTEROCOCCI_IND;
                bcLandSampleNew.SR_ENTEROCOCCI = SR_ENTEROCOCCI;
                bcLandSampleNew.SR_FLOW = SR_FLOW;
                bcLandSampleNew.SR_SAMPLE_AGENCY = SR_SAMPLE_AGENCY;
                bcLandSampleNew.SR_OLD_KEY = SR_OLD_KEY;
                bcLandSampleNew.SR_SAMPLE_DEPTH = SR_SAMPLE_DEPTH;
                bcLandSampleNew.SR_SALINITY = SR_SALINITY;
                bcLandSampleNew.SR_TEMPERATURE = SR_TEMPERATURE;
                bcLandSampleNew.SR_TIDE_CODE = SR_TIDE_CODE;
                bcLandSampleNew.SR_OBS = SR_OBS;

                bcLandSampleList.Add(bcLandSampleNew);
            }

            reader.Close();
            conn.Close();

            int InDBAlready = 0;
            int skip = 0;

            lblTotal.Text = (bcLandSampleList.Count() + InDBAlready).ToString();
            int take = bcLandSampleList.Count() / DoBWCount;
            for (int i = 0; i < DoBWCount; i++)
            {
                BWBCLandSample bwBCLandSample = new BWBCLandSample();
                bwBCLandSample.bcLandSampleList = bcLandSampleList;
                bwBCLandSample.skip = skip;
                if (i == (DoBWCount - 1))
                {
                    take = 1000000;
                }
                bwBCLandSample.take = take;
                bwListBCLandSample[i].RunWorkerAsync(bwBCLandSample);
                skip += take;
            }

            return;
        }
        private void LoadBCLand_Based_Sample_Stations()
        {
            List<BCLandSampleStation> bcLandSampleStationList = new List<BCLandSampleStation>();

            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OleDb.12.0;Data Source=C:\CSSP\BC_Data\CSSP_Pacific_WQ_Data_20180219.accdb");
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Land_Based_Sample_Station]", conn);
            OleDbDataReader reader = cmd.ExecuteReader();

            List<string> FieldNameList = new List<string>();

            FieldNameList = new List<string>() { "OID", "SS_STATION_CODE",
                "SS_DESCRIPTION", "SS_LATITUDE_DEGREES", "SS_LATITUDE_MINUTES", "SS_LONGITUDE_DEGREES", "SS_LONGITUDE_MINUTES",
                "SS_SHELLFISH_SECTOR", "SS_DFO_SUBAREA", "SS_HARVEST_TYPE", "SS_STATUS",
                "SS_REMOTE_STATUS", "SS_LASTUPDATE", "SS_ENTEREDBY", "LAT", "LON", "METADATA", "DATALINK", "SS_STATION_TYPE", "OLD_KEY", "OLD_KEY2", "QA" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            reader = cmd.ExecuteReader();

            int Count = 0;
            while (reader.Read())
            {
                Count += 1;
                Application.DoEvents();

                int? OID = -1;
                string SS_STATION_CODE = "";
                string SS_DESCRIPTION = "";
                float? SS_LATITUDE_DEGREES = -1f;
                float? SS_LATITUDE_MINUTES = -1f;
                float? SS_LONGITUDE_DEGREES = -1f;
                float? SS_LONGITUDE_MINUTES = -1f;
                string SS_SHELLFISH_SECTOR = "";
                string SS_DFO_SUBAREA = "";
                string SS_HARVEST_TYPE = "";
                string SS_STATUS = "";
                string SS_REMOTE_STATUS = "";
                DateTime SS_LASTUPDATE = new DateTime(2050, 1, 1);
                string SS_ENTEREDBY = "";
                float? LAT = -1f;
                float? LON = -1f;
                string METADATA = "";
                string DATALINK = "";
                string SS_STATION_TYPE = "";
                string OLD_KEY = "";
                string OLD_KEY2 = "";
                string QA = "";

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    OID = null;
                }
                else
                {
                    OID = (int)(reader.GetValue(0));
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    SS_STATION_CODE = null;
                }
                else
                {
                    SS_STATION_CODE = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    SS_DESCRIPTION = null;
                }
                else
                {
                    SS_DESCRIPTION = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    SS_LATITUDE_DEGREES = null;
                }
                else
                {
                    SS_LATITUDE_DEGREES = (float)(byte)(reader.GetValue(3));
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    SS_LATITUDE_MINUTES = null;
                }
                else
                {
                    SS_LATITUDE_MINUTES = (float)(reader.GetValue(4));
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    SS_LONGITUDE_DEGREES = null;
                }
                else
                {
                    SS_LONGITUDE_DEGREES = (float)(byte)reader.GetValue(5);
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    SS_LONGITUDE_MINUTES = null;
                }
                else
                {
                    SS_LONGITUDE_MINUTES = (float)(reader.GetValue(6));
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    SS_SHELLFISH_SECTOR = null;
                }
                else
                {
                    SS_SHELLFISH_SECTOR = reader.GetValue(7).ToString().Trim();
                }

                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    SS_DFO_SUBAREA = null;
                }
                else
                {
                    SS_DFO_SUBAREA = reader.GetValue(8).ToString().Trim();
                }

                if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                {
                    SS_HARVEST_TYPE = null;
                }
                else
                {
                    SS_HARVEST_TYPE = reader.GetValue(9).ToString().Trim();
                }

                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                {
                    SS_STATUS = null;
                }
                else
                {
                    SS_STATUS = reader.GetValue(10).ToString().Trim();
                }

                if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                {
                    SS_REMOTE_STATUS = null;
                }
                else
                {
                    SS_REMOTE_STATUS = reader.GetValue(11).ToString().Trim();
                }

                if (reader.GetValue(12).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(12).ToString()))
                {
                    SS_LASTUPDATE = new DateTime(2050, 1, 1);
                }
                else
                {
                    SS_LASTUPDATE = (DateTime)(reader.GetValue(12));
                }

                if (reader.GetValue(13).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(13).ToString()))
                {
                    SS_ENTEREDBY = null;
                }
                else
                {
                    SS_ENTEREDBY = reader.GetValue(13).ToString().Trim();
                }

                if (reader.GetValue(14).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(14).ToString()))
                {
                    LAT = null;
                }
                else
                {
                    LAT = (float)(double)reader.GetValue(14);
                }

                if (reader.GetValue(15).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(15).ToString()))
                {
                    LON = null;
                }
                else
                {
                    LON = (float)(double)reader.GetValue(15);
                }

                if (reader.GetValue(16).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(16).ToString()))
                {
                    METADATA = null;
                }
                else
                {
                    METADATA = reader.GetValue(16).ToString().Trim();
                }

                if (reader.GetValue(17).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(17).ToString()))
                {
                    DATALINK = null;
                }
                else
                {
                    DATALINK = reader.GetValue(17).ToString().Trim();
                }

                if (reader.GetValue(18).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(18).ToString()))
                {
                    SS_STATION_TYPE = null;
                }
                else
                {
                    SS_STATION_TYPE = reader.GetValue(18).ToString().Trim();
                }

                if (reader.GetValue(19).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(19).ToString()))
                {
                    OLD_KEY = null;
                }
                else
                {
                    OLD_KEY = reader.GetValue(19).ToString().Trim();
                }

                if (reader.GetValue(20).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(20).ToString()))
                {
                    DATALINK = null;
                }
                else
                {
                    OLD_KEY2 = reader.GetValue(20).ToString().Trim();
                }

                if (reader.GetValue(21).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(21).ToString()))
                {
                    QA = null;
                }
                else
                {
                    QA = reader.GetValue(21).ToString().Trim();
                }


                BCLandSampleStation bcLandSampleStationNew = new BCLandSampleStation();
                bcLandSampleStationNew.OID = OID;
                bcLandSampleStationNew.SS_STATION_CODE = SS_STATION_CODE;
                bcLandSampleStationNew.SS_DESCRIPTION = SS_DESCRIPTION;
                bcLandSampleStationNew.SS_LATITUDE_DEGREES = SS_LATITUDE_DEGREES;
                bcLandSampleStationNew.SS_LATITUDE_MINUTES = SS_LATITUDE_MINUTES;
                bcLandSampleStationNew.SS_LONGITUDE_DEGREES = SS_LONGITUDE_DEGREES;
                bcLandSampleStationNew.SS_LONGITUDE_MINUTES = SS_LONGITUDE_MINUTES;
                bcLandSampleStationNew.SS_SHELLFISH_SECTOR = SS_SHELLFISH_SECTOR;
                bcLandSampleStationNew.SS_DFO_SUBAREA = SS_DFO_SUBAREA;
                bcLandSampleStationNew.SS_HARVEST_TYPE = SS_HARVEST_TYPE;
                bcLandSampleStationNew.SS_STATUS = SS_STATUS;
                bcLandSampleStationNew.SS_REMOTE_STATUS = SS_REMOTE_STATUS;
                bcLandSampleStationNew.SS_LASTUPDATE = SS_LASTUPDATE;
                bcLandSampleStationNew.SS_ENTEREDBY = SS_ENTEREDBY;
                bcLandSampleStationNew.LAT = LAT;
                bcLandSampleStationNew.LON = LON;
                bcLandSampleStationNew.METADATA = METADATA;
                bcLandSampleStationNew.DATALINK = DATALINK;
                bcLandSampleStationNew.SS_STATION_TYPE = SS_STATION_TYPE;
                bcLandSampleStationNew.OLD_KEY = OLD_KEY;
                bcLandSampleStationNew.OLD_KEY2 = OLD_KEY2;
                bcLandSampleStationNew.QA = QA;

                bcLandSampleStationList.Add(bcLandSampleStationNew);
            }

            reader.Close();
            conn.Close();

            int InDBAlready = 0;
            int skip = 0;

            lblTotal.Text = (bcLandSampleStationList.Count() + InDBAlready).ToString();
            int take = bcLandSampleStationList.Count() / DoBWCount;
            for (int i = 0; i < DoBWCount; i++)
            {
                BWBCLandStation bwBCLandStation = new BWBCLandStation();
                bwBCLandStation.bcLandStationList = bcLandSampleStationList;
                bwBCLandStation.skip = skip;
                if (i == (DoBWCount - 1))
                {
                    take = 1000000;
                }
                bwBCLandStation.take = take;
                bwListBCLandStation[i].RunWorkerAsync(bwBCLandStation);
                skip += take;
            }

            return;
        }
        private void LoadBCMarine_Sample_Readings()
        {
            List<BCMarineSample> bcMarineSampleList = new List<BCMarineSample>();

            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OleDb.12.0;Data Source=C:\CSSP\BC_Data\CSSP_Pacific_WQ_Data_20180219.accdb");
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Marine Sample Reading]", conn);
            OleDbDataReader reader = cmd.ExecuteReader();

            List<string> FieldNameList = new List<string>();

            FieldNameList = new List<string>() { "OID", "SR_SURVEY", "SR_STATION_CODE",
                "SR_ANALYSIS_TYPE", "SR_READING_DATE", "SR_READING_TIME", "SR_SAMPLE_TYPE", "SR_SAMPLE_DEPTH",
                "SR_FECAL_COLIFORM_IND", "SR_FECAL_COLIFORM", "SR_ENTEROCOCCI_IND", "SR_ENTEROCOCCI",
                "SR_SALINITY", "SR_TEMPERATURE", "SR_TURBIDITY", "SR_TIDE_CODE", "SR_SPECIES", "SR_SAMPLE_AGENCY", "SR_OBS" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            reader = cmd.ExecuteReader();

            int Count = 0;
            while (reader.Read())
            {
                Count += 1;
                Application.DoEvents();

                int? OID = -1;
                int? SR_SURVEY = -1;
                string SR_STATION_CODE = "";
                string SR_ANALYSIS_TYPE = "";
                DateTime SR_READING_DATE = new DateTime(2050, 1, 1);
                string SR_READING_TIME = "";
                string SR_SAMPLE_TYPE = "";
                float? SR_SAMPLE_DEPTH = -1;
                string SR_FECAL_COLIFORM_IND = "";
                int? SR_FECAL_COLIFORM = -1;
                string SR_ENTEROCOCCI_IND = "";
                int? SR_ENTEROCOCCI = -1;
                float? SR_SALINITY = -1;
                float? SR_TEMPERATURE = -1;
                float? SR_TURBIDITY = -1;
                string SR_TIDE_CODE = "";
                string SR_SPECIES = "";
                int? SR_SAMPLE_AGENCY = -1;
                string SR_OBS = "";

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    OID = null;
                }
                else
                {
                    OID = (int)(reader.GetValue(0));
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    SR_SURVEY = null;
                }
                else
                {
                    SR_SURVEY = (int)(Int16)(reader.GetValue(1));
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    SR_STATION_CODE = null;
                }
                else
                {
                    SR_STATION_CODE = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    SR_ANALYSIS_TYPE = null;
                }
                else
                {
                    SR_ANALYSIS_TYPE = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    SR_READING_DATE = new DateTime(2030, 1, 1);
                }
                else
                {
                    SR_READING_DATE = (DateTime)(reader.GetValue(4));
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    SR_READING_TIME = null;
                }
                else
                {
                    SR_READING_TIME = reader.GetValue(5).ToString().Trim();
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    SR_SAMPLE_TYPE = null;
                }
                else
                {
                    SR_SAMPLE_TYPE = reader.GetValue(6).ToString().Trim();
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    SR_SAMPLE_DEPTH = null;
                }
                else
                {
                    SR_SAMPLE_DEPTH = (float)(reader.GetValue(7));
                }

                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    SR_FECAL_COLIFORM_IND = null;
                }
                else
                {
                    SR_FECAL_COLIFORM_IND = reader.GetValue(8).ToString().Trim();
                }

                if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                {
                    SR_FECAL_COLIFORM = null;
                }
                else
                {
                    SR_FECAL_COLIFORM = (int)(reader.GetValue(9));
                }

                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                {
                    SR_ENTEROCOCCI_IND = null;
                }
                else
                {
                    SR_ENTEROCOCCI_IND = reader.GetValue(10).ToString().Trim();
                }

                if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                {
                    SR_ENTEROCOCCI = null;
                }
                else
                {
                    SR_ENTEROCOCCI = (int)(reader.GetValue(11));
                }

                if (reader.GetValue(12).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(12).ToString()))
                {
                    SR_SALINITY = null;
                }
                else
                {
                    SR_SALINITY = (float)(reader.GetValue(12));
                }

                if (reader.GetValue(13).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(13).ToString()))
                {
                    SR_TEMPERATURE = null;
                }
                else
                {
                    SR_TEMPERATURE = (float)(reader.GetValue(13));
                }

                if (reader.GetValue(14).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(14).ToString()))
                {
                    SR_TURBIDITY = null;
                }
                else
                {
                    SR_TURBIDITY = (float)(reader.GetValue(14));
                }

                if (reader.GetValue(15).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(15).ToString()))
                {
                    SR_TIDE_CODE = null;
                }
                else
                {
                    SR_TIDE_CODE = reader.GetValue(15).ToString().Trim();
                }

                if (reader.GetValue(16).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(16).ToString()))
                {
                    SR_SPECIES = null;
                }
                else
                {
                    SR_SPECIES = reader.GetValue(16).ToString().Trim();
                }

                if (reader.GetValue(17).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(17).ToString()))
                {
                    SR_SAMPLE_AGENCY = null;
                }
                else
                {
                    SR_SAMPLE_AGENCY = (int)(Int16)(reader.GetValue(17));
                }

                if (reader.GetValue(18).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(18).ToString()))
                {
                    SR_OBS = null;
                }
                else
                {
                    SR_OBS = reader.GetValue(18).ToString().Trim();
                }


                BCMarineSample bcMarineSampleNew = new BCMarineSample();
                bcMarineSampleNew.SR_SURVEY = SR_SURVEY;
                bcMarineSampleNew.SR_STATION_CODE = SR_STATION_CODE;
                bcMarineSampleNew.SR_ANALYSIS_TYPE = SR_ANALYSIS_TYPE;
                bcMarineSampleNew.SR_READING_DATE = SR_READING_DATE;
                bcMarineSampleNew.SR_READING_TIME = SR_READING_TIME;
                bcMarineSampleNew.SR_SAMPLE_TYPE = SR_SAMPLE_TYPE;
                bcMarineSampleNew.SR_SAMPLE_DEPTH = SR_SAMPLE_DEPTH;
                bcMarineSampleNew.SR_FECAL_COLIFORM_IND = SR_FECAL_COLIFORM_IND;
                bcMarineSampleNew.SR_FECAL_COLIFORM = SR_FECAL_COLIFORM;
                bcMarineSampleNew.SR_ENTEROCOCCI_IND = SR_ENTEROCOCCI_IND;
                bcMarineSampleNew.SR_ENTEROCOCCI = SR_ENTEROCOCCI;
                bcMarineSampleNew.SR_SALINITY = SR_SALINITY;
                bcMarineSampleNew.SR_TEMPERATURE = SR_TEMPERATURE;
                bcMarineSampleNew.SR_TURBIDITY = SR_TURBIDITY;
                bcMarineSampleNew.SR_TIDE_CODE = SR_TIDE_CODE;
                bcMarineSampleNew.SR_SPECIES = SR_SPECIES;
                bcMarineSampleNew.SR_SAMPLE_AGENCY = SR_SAMPLE_AGENCY;
                bcMarineSampleNew.SR_OBS = SR_OBS;

                bcMarineSampleList.Add(bcMarineSampleNew);
            }

            reader.Close();
            conn.Close();

            int InDBAlready = 0;
            int skip = 0;

            lblTotal.Text = (bcMarineSampleList.Count() + InDBAlready).ToString();
            int take = bcMarineSampleList.Count() / DoBWCount;
            for (int i = 0; i < DoBWCount; i++)
            {
                BWBCMarineSample bwBCMarineSample = new BWBCMarineSample();
                bwBCMarineSample.bcMarineSampleList = bcMarineSampleList;
                bwBCMarineSample.skip = skip;
                if (i == (DoBWCount - 1))
                {
                    take = 1000000;
                }
                bwBCMarineSample.take = take;
                bwListBCMarineSample[i].RunWorkerAsync(bwBCMarineSample);
                skip += take;
            }

            return;
        }
        private void LoadBCMarine_Sample_Stations()
        {
            List<BCMarineSampleStation> bcMarineSampleStationList = new List<BCMarineSampleStation>();

            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OleDb.12.0;Data Source=C:\CSSP\BC_Data\CSSP_Pacific_WQ_Data_20180219.accdb");
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Marine_Sample_Station]", conn);
            OleDbDataReader reader = cmd.ExecuteReader();

            List<string> FieldNameList = new List<string>();

            FieldNameList = new List<string>() { "OID", "SS_STATION_CODE",
                "SS_DESCRIPTION", "SS_LATITUDE_DEGREES", "SS_LATITUDE_MINUTES", "SS_LONGITUDE_DEGREES", "SS_LONGITUDE_MINUTES",
                "SS_SHELLFISH_SECTOR", "SS_DFO_SUBAREA", "SS_HARVEST_TYPE", "SS_STATUS",
                "SS_REMOTE_STATUS", "SS_LASTUPDATE", "SS_ENTEREDBY", "LAT", "LON", "METADATA", "DATALINK" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            reader = cmd.ExecuteReader();

            int Count = 0;
            while (reader.Read())
            {
                Count += 1;
                Application.DoEvents();

                int? OID = -1;
                string SS_STATION_CODE = "";
                string SS_DESCRIPTION = "";
                float? SS_LATITUDE_DEGREES = -1f;
                float? SS_LATITUDE_MINUTES = -1f;
                float? SS_LONGITUDE_DEGREES = -1f;
                float? SS_LONGITUDE_MINUTES = -1f;
                string SS_SHELLFISH_SECTOR = "";
                string SS_DFO_SUBAREA = "";
                string SS_HARVEST_TYPE = "";
                string SS_STATUS = "";
                string SS_REMOTE_STATUS = "";
                DateTime SS_LASTUPDATE = new DateTime(2050, 1, 1);
                string SS_ENTEREDBY = "";
                float? LAT = -1f;
                float? LON = -1f;
                string METADATA = "";
                string DATALINK = "";

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    OID = null;
                }
                else
                {
                    OID = (int)(reader.GetValue(0));
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    SS_STATION_CODE = null;
                }
                else
                {
                    SS_STATION_CODE = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    SS_DESCRIPTION = null;
                }
                else
                {
                    SS_DESCRIPTION = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    SS_LATITUDE_DEGREES = null;
                }
                else
                {
                    SS_LATITUDE_DEGREES = (float)(double)(reader.GetValue(3));
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    SS_LATITUDE_MINUTES = null;
                }
                else
                {
                    SS_LATITUDE_MINUTES = (float)(double)reader.GetValue(4);
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    SS_LONGITUDE_DEGREES = null;
                }
                else
                {
                    SS_LONGITUDE_DEGREES = (float)(double)reader.GetValue(5);
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    SS_LONGITUDE_MINUTES = null;
                }
                else
                {
                    SS_LONGITUDE_MINUTES = (float)(double)(reader.GetValue(6));
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    SS_SHELLFISH_SECTOR = null;
                }
                else
                {
                    SS_SHELLFISH_SECTOR = reader.GetValue(7).ToString().Trim();
                }

                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    SS_DFO_SUBAREA = null;
                }
                else
                {
                    SS_DFO_SUBAREA = reader.GetValue(8).ToString().Trim();
                }

                if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                {
                    SS_HARVEST_TYPE = null;
                }
                else
                {
                    SS_HARVEST_TYPE = reader.GetValue(9).ToString().Trim();
                }

                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                {
                    SS_STATUS = null;
                }
                else
                {
                    SS_STATUS = reader.GetValue(10).ToString().Trim();
                }

                if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                {
                    SS_REMOTE_STATUS = null;
                }
                else
                {
                    SS_REMOTE_STATUS = reader.GetValue(11).ToString().Trim();
                }

                if (reader.GetValue(12).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(12).ToString()))
                {
                    SS_LASTUPDATE = new DateTime(2050, 1, 1);
                }
                else
                {
                    SS_LASTUPDATE = (DateTime)(reader.GetValue(12));
                }

                if (reader.GetValue(13).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(13).ToString()))
                {
                    SS_ENTEREDBY = null;
                }
                else
                {
                    SS_ENTEREDBY = reader.GetValue(13).ToString().Trim();
                }

                if (reader.GetValue(14).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(14).ToString()))
                {
                    LAT = null;
                }
                else
                {
                    LAT = (float)(double)reader.GetValue(14);
                }

                if (reader.GetValue(15).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(15).ToString()))
                {
                    LON = null;
                }
                else
                {
                    LON = (float)(double)reader.GetValue(15);
                }

                if (reader.GetValue(16).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(16).ToString()))
                {
                    METADATA = null;
                }
                else
                {
                    METADATA = reader.GetValue(16).ToString().Trim();
                }

                if (reader.GetValue(17).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(17).ToString()))
                {
                    DATALINK = null;
                }
                else
                {
                    DATALINK = reader.GetValue(17).ToString().Trim();
                }


                BCMarineSampleStation bcMarineSampleStationNew = new BCMarineSampleStation();
                bcMarineSampleStationNew.OID = OID;
                bcMarineSampleStationNew.SS_STATION_CODE = SS_STATION_CODE;
                bcMarineSampleStationNew.SS_DESCRIPTION = SS_DESCRIPTION;
                bcMarineSampleStationNew.SS_LATITUDE_DEGREES = SS_LATITUDE_DEGREES;
                bcMarineSampleStationNew.SS_LATITUDE_MINUTES = SS_LATITUDE_MINUTES;
                bcMarineSampleStationNew.SS_LONGITUDE_DEGREES = SS_LONGITUDE_DEGREES;
                bcMarineSampleStationNew.SS_LONGITUDE_MINUTES = SS_LONGITUDE_MINUTES;
                bcMarineSampleStationNew.SS_SHELLFISH_SECTOR = SS_SHELLFISH_SECTOR;
                bcMarineSampleStationNew.SS_DFO_SUBAREA = SS_DFO_SUBAREA;
                bcMarineSampleStationNew.SS_HARVEST_TYPE = SS_HARVEST_TYPE;
                bcMarineSampleStationNew.SS_STATUS = SS_STATUS;
                bcMarineSampleStationNew.SS_REMOTE_STATUS = SS_REMOTE_STATUS;
                bcMarineSampleStationNew.SS_LASTUPDATE = SS_LASTUPDATE;
                bcMarineSampleStationNew.SS_ENTEREDBY = SS_ENTEREDBY;
                bcMarineSampleStationNew.LAT = LAT;
                bcMarineSampleStationNew.LON = LON;
                bcMarineSampleStationNew.METADATA = METADATA;
                bcMarineSampleStationNew.DATALINK = DATALINK;

                bcMarineSampleStationList.Add(bcMarineSampleStationNew);
            }

            reader.Close();
            conn.Close();

            int InDBAlready = 0;
            int skip = 0;

            lblTotal.Text = (bcMarineSampleStationList.Count() + InDBAlready).ToString();
            int take = bcMarineSampleStationList.Count() / DoBWCount;
            for (int i = 0; i < DoBWCount; i++)
            {
                BWBCMarineStation bwBCMarineStation = new BWBCMarineStation();
                bwBCMarineStation.bcMarineStationList = bcMarineSampleStationList;
                bwBCMarineStation.skip = skip;
                if (i == (DoBWCount - 1))
                {
                    take = 1000000;
                }
                bwBCMarineStation.take = take;
                bwListBCMarineStation[i].RunWorkerAsync(bwBCMarineStation);
                skip += take;
            }

            return;
        }
        private void LoadBCPolSource()
        {
            List<BCPolSource> bcPolSourceList = new List<BCPolSource>();

            FileInfo fi = new FileInfo(@"C:\BC_Data\PSI_20130626.xls");

            string connectionString = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + fi.FullName + ";Extended Properties=Excel 8.0;";

            Application.DoEvents();

            OleDbCommand comm = new OleDbCommand("Select * from [PSI_20130626$];");
            OleDbConnection conn = new OleDbConnection(connectionString);

            conn.Open();
            OleDbDataReader reader;
            comm.Connection = conn;
            reader = comm.ExecuteReader();

            List<string> FieldNameList = new List<string>();

            FieldNameList = new List<string>() { "OBJECTID", "Verified", "Source_nme", "Label", "Name", "Location", "LCode",
                "ICode", "Key_", "X_calc", "Y_calc", "DATALINK", "yyyymmdd", "Remarks", "LatDegMin", "LonDegMin" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            reader = comm.ExecuteReader();

            int Count = 0;
            while (reader.Read())
            {
                Count += 1;

                int? OBJECTID = 0;
                string Verified = "";
                string Source_nme = "";
                string Label = "";
                string Name = "";
                string Location = "";
                string LCode = "";
                string ICode = "";
                string Key_ = "";
                float? X_calc = 0.0f;
                float? Y_calc = 0.0f;
                string DATALINK = "";
                DateTime yyyymmdd = new DateTime(2050, 1, 1);
                string Remarks = "";
                string LatDegMin = "";
                string LonDegMin = "";

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString().Trim()))
                {
                    OBJECTID = null;
                }
                else
                {
                    OBJECTID = (int)((double)reader.GetValue(0));
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString().Trim()))
                {
                    Verified = null;
                }
                else
                {
                    Verified = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString().Trim()))
                {
                    Source_nme = null;
                }
                else
                {
                    Source_nme = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString().Trim()))
                {
                    Label = null;
                }
                else
                {
                    Label = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString().Trim()))
                {
                    Name = null;
                }
                else
                {
                    Name = reader.GetValue(4).ToString().Trim();
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString().Trim()))
                {
                    Location = null;
                }
                else
                {
                    Location = reader.GetValue(5).ToString().Trim();
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString().Trim()))
                {
                    LCode = null;
                }
                else
                {
                    LCode = reader.GetValue(6).ToString().Trim();
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString().Trim()))
                {
                    ICode = null;
                }
                else
                {
                    ICode = reader.GetValue(7).ToString().Trim();
                }

                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString().Trim()))
                {
                    Key_ = null;
                }
                else
                {
                    Key_ = reader.GetValue(8).ToString().Trim();
                }

                if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString().Trim()))
                {
                    X_calc = null;
                }
                else
                {
                    X_calc = (float)(double)(reader.GetValue(9));
                }

                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString().Trim()))
                {
                    Y_calc = null;
                }
                else
                {
                    Y_calc = (float)(double)(reader.GetValue(10));
                }

                if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString().Trim()))
                {
                    DATALINK = null;
                }
                else
                {
                    DATALINK = reader.GetValue(11).ToString().Trim();
                }

                if (reader.GetValue(12).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(12).ToString().Trim()))
                {
                    yyyymmdd = new DateTime(2050, 1, 1);
                }
                else
                {
                    int year = int.Parse(reader.GetValue(12).ToString().Substring(0, 4));
                    int month = int.Parse(reader.GetValue(12).ToString().Substring(4, 2));
                    int day = int.Parse(reader.GetValue(12).ToString().Substring(6, 2));
                    if (year == 2012 && month == 2 && day == 30)
                    {
                        year = 2012;
                        month = 3;
                        day = 1;
                    }
                    if (year == 2011 && month == 0 && day == 7)
                    {
                        year = 2011;
                        month = 1;
                        day = 7;
                    }
                    yyyymmdd = new DateTime(year, month, day);
                }

                if (reader.GetValue(13).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(13).ToString().Trim()))
                {
                    Remarks = null;
                }
                else
                {
                    Remarks = reader.GetValue(13).ToString().Trim();
                }

                if (reader.GetValue(14).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(14).ToString().Trim()))
                {
                    LatDegMin = null;
                }
                else
                {
                    LatDegMin = reader.GetValue(14).ToString().Trim();
                }

                if (reader.GetValue(15).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(15).ToString().Trim()))
                {
                    LonDegMin = null;
                }
                else
                {
                    LonDegMin = reader.GetValue(15).ToString().Trim();
                }


                BCPolSource bcPolSourceNew = new BCPolSource();
                bcPolSourceNew.OBJECTID = OBJECTID;
                bcPolSourceNew.Verified = Verified;
                bcPolSourceNew.Source_nme = Source_nme;
                bcPolSourceNew.Label = Label;
                bcPolSourceNew.Name = Name;
                bcPolSourceNew.Location = Location;
                bcPolSourceNew.LCode = LCode;
                bcPolSourceNew.ICode = ICode;
                bcPolSourceNew.Key_ = Key_;
                bcPolSourceNew.X_calc = X_calc;
                bcPolSourceNew.Y_calc = Y_calc;
                bcPolSourceNew.DATALINK = DATALINK;
                bcPolSourceNew.yyyymmdd = yyyymmdd;
                bcPolSourceNew.Remarks = Remarks;
                bcPolSourceNew.LatDegMin = LatDegMin;
                bcPolSourceNew.LonDegMin = LonDegMin;

                bcPolSourceList.Add(bcPolSourceNew);
            }

            int InDBAlready = 0;
            int skip = 0;

            lblTotal.Text = (bcPolSourceList.Count() + InDBAlready).ToString();
            int take = bcPolSourceList.Count() / DoBWCount;
            for (int i = 0; i < DoBWCount; i++)
            {
                BWBCPolSource bwBCPolSource = new BWBCPolSource();
                bwBCPolSource.bcPolSourceList = bcPolSourceList;
                bwBCPolSource.skip = skip;
                if (i == (DoBWCount - 1))
                {
                    take = 1000000;
                }
                bwBCPolSource.take = take;
                bwListBCPolSource[i].RunWorkerAsync(bwBCPolSource);
                skip += take;
            }

            return;

        }
        private void LoadBCSurveys()
        {
            List<BCSurvey> bcSurveyList = new List<BCSurvey>();

            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OleDb.12.0;Data Source=C:\CSSP\BC_Data\CSSP_Pacific_WQ_Data_20180219.accdb");
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Survey]", conn);
            OleDbDataReader reader = cmd.ExecuteReader();

            List<string> FieldNameList = new List<string>();

            FieldNameList = new List<string>() { "S_ID_NUMBER", "S_SHELLFISH_SECTOR", "S_DESCRIPTION", "S_SURVEY_TYPE",
                "S_START_DATE", "S_END_DATE", "S_COMMENT" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            reader = cmd.ExecuteReader();

            int Count = 0;
            while (reader.Read())
            {
                Count += 1;
                Application.DoEvents();

                int? S_ID_NUMBER = 0;
                string S_SHELLFISH_SECTOR = "";
                string S_DESCRIPTION = "";
                string S_SURVEY_TYPE = "";
                DateTime S_START_DATE = new DateTime(2050, 1, 1);
                DateTime S_END_DATE = new DateTime(2050, 1, 1);
                string S_COMMENT = "";

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    S_ID_NUMBER = null;
                }
                else
                {
                    S_ID_NUMBER = (int)(Int16)(reader.GetValue(0));
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    S_SHELLFISH_SECTOR = null;
                }
                else
                {
                    S_SHELLFISH_SECTOR = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    S_DESCRIPTION = null;
                }
                else
                {
                    S_DESCRIPTION = reader.GetValue(2).ToString().Trim();
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    S_SURVEY_TYPE = null;
                }
                else
                {
                    S_SURVEY_TYPE = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    S_START_DATE = new DateTime(2050, 1, 1);
                }
                else
                {
                    S_START_DATE = (DateTime)(reader.GetValue(4));
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    S_END_DATE = new DateTime(2050, 1, 1);
                }
                else
                {
                    S_END_DATE = (DateTime)(reader.GetValue(5));
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    S_COMMENT = null;
                }
                else
                {
                    S_COMMENT = reader.GetValue(6).ToString().Trim();
                }

                BCSurvey bcSurveyNew = new BCSurvey();
                bcSurveyNew.S_ID_NUMBER = S_ID_NUMBER;
                bcSurveyNew.S_SHELLFISH_SECTOR = S_SHELLFISH_SECTOR;
                bcSurveyNew.S_DESCRIPTION = S_DESCRIPTION;
                bcSurveyNew.S_SURVEY_TYPE = S_SURVEY_TYPE;
                bcSurveyNew.S_START_DATE = S_START_DATE;
                bcSurveyNew.S_END_DATE = S_END_DATE;
                bcSurveyNew.S_COMMENT = S_COMMENT;

                bcSurveyList.Add(bcSurveyNew);
            }

            reader.Close();
            conn.Close();

            int InDBAlready = 0;
            int skip = 0;

            lblTotal.Text = (bcSurveyList.Count() + InDBAlready).ToString();
            int take = bcSurveyList.Count() / DoBWCount;
            for (int i = 0; i < DoBWCount; i++)
            {
                BWBCRun bwBCRun = new BWBCRun();
                bwBCRun.bcRunList = bcSurveyList;
                bwBCRun.skip = skip;
                if (i == (DoBWCount - 1))
                {
                    take = 1000000;
                }
                bwBCRun.take = take;
                bwListBCRun[i].RunWorkerAsync(bwBCRun);
                skip += take;
            }

            return;
        }
        private void LoadBCWeather_Readings()
        {
            List<BCClimateSite> bcClimateSiteList = new List<BCClimateSite>();

            // reading all distinct weather station
            FileInfo fi = new FileInfo(@"C:\BC_Data\ClimateSiteBC.xls");

            string connectionString = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + fi.FullName + ";Extended Properties=Excel 8.0;";

            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("Select * from [ClimateSiteBC$]", conn);
            OleDbDataReader reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                string BCText = "";
                string ClimateSiteName = "";
                string ClimateID = "";

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    BCText = null;
                }
                else
                {
                    BCText = reader.GetValue(0).ToString().Trim().ToUpper();
                }


                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    ClimateSiteName = null;
                }
                else
                {
                    ClimateSiteName = reader.GetValue(1).ToString().Trim();
                }

                if (!string.IsNullOrEmpty(BCText) && (BCText == "HARDY ISLAND" || BCText == "HARDY ISLAND(CON.)" || BCText == "HARDY ISLAND (CON.)"))
                {
                    ClimateID = "104LC3N";
                }
                else if (!string.IsNullOrEmpty(BCText) && (BCText == "SEWALL (QCI)" || BCText == "SEWALL, Q.C.I."))
                {
                    ClimateID = "105PA91";
                }
                else if (!string.IsNullOrEmpty(BCText) && (BCText == "MORESBY ISLAND MITCHELL INLET"))
                {
                    ClimateID = "10551R8";
                }
                else
                {
                    if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                    {
                        ClimateID = null;
                    }
                    else
                    {
                        ClimateID = reader.GetValue(2).ToString().Trim().ToUpper();
                    }
                }

                if (BCText != null)
                {
                    BCClimateSite bcClimateSiteNew = new BCClimateSite();
                    bcClimateSiteNew.BCText = BCText;
                    bcClimateSiteNew.ClimateSiteName = ClimateSiteName;
                    bcClimateSiteNew.ClimateID = ClimateID;

                    bcClimateSiteList.Add(bcClimateSiteNew);
                }
                Application.DoEvents();

            }
            reader.Close();
            conn.Close();


            List<BCPrecipitation> bcPrecList = new List<BCPrecipitation>();

            conn = new OleDbConnection(@"Provider=Microsoft.ACE.OleDb.12.0;Data Source=C:\CSSP\BC_Data\CSSP_Pacific_WQ_Data_20180219.accdb");
            conn.Open();

            cmd = new OleDbCommand("SELECT * FROM [Weather Reading]", conn);
            reader = cmd.ExecuteReader();

            List<string> FieldNameList = new List<string>();

            FieldNameList = new List<string>() { "WR_DATE", "WR_SURVEY", "WR_PRECIPITATION_RAIN", "WR_PRECIPITATION_SNOW",
                "WR_TOTAL_PRECIPITATION", "WR_RECORDING_STATION" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();
            reader = cmd.ExecuteReader();

            int Count = 0;
            while (reader.Read())
            {
                Count += 1;
                Application.DoEvents();

                DateTime WR_DATE = new DateTime(2050, 1, 1);
                int? WR_SURVEY = 0;
                float? WR_PRECIPITATION_RAIN = 999;
                float? WR_PRECIPITATION_SNOW = 999;
                float? WR_TOTAL_PRECIPITATION = 999;
                string WR_RECORDING_STATION = "";

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    WR_DATE = new DateTime(2050, 1, 1);
                }
                else
                {
                    WR_DATE = (DateTime)(reader.GetValue(0));
                }

                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    WR_SURVEY = null;
                }
                else
                {
                    WR_SURVEY = (int)(Int16)(reader.GetValue(1));
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    WR_PRECIPITATION_RAIN = null;
                }
                else
                {
                    WR_PRECIPITATION_RAIN = (float)(reader.GetValue(2));
                }

                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    WR_PRECIPITATION_SNOW = null;
                }
                else
                {
                    WR_PRECIPITATION_SNOW = (float)(reader.GetValue(3));
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    WR_TOTAL_PRECIPITATION = null;
                }
                else
                {
                    WR_TOTAL_PRECIPITATION = (float)(reader.GetValue(4));
                }

                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    WR_RECORDING_STATION = null;
                }
                else
                {
                    WR_RECORDING_STATION = reader.GetValue(5).ToString().Trim();
                    if (WR_RECORDING_STATION.Substring(0, 1) == "-")
                    {
                        WR_RECORDING_STATION = WR_RECORDING_STATION.Substring(1).Trim();
                    }
                    WR_RECORDING_STATION = WR_RECORDING_STATION.ToUpper();
                }

                BCPrecipitation bcPrecNew = new BCPrecipitation();
                bcPrecNew.WR_DATE = WR_DATE;
                bcPrecNew.WR_SURVEY = WR_SURVEY;
                bcPrecNew.WR_PRECIPITATION_RAIN = WR_PRECIPITATION_RAIN;
                bcPrecNew.WR_PRECIPITATION_SNOW = WR_PRECIPITATION_SNOW;
                bcPrecNew.WR_TOTAL_PRECIPITATION = WR_TOTAL_PRECIPITATION;
                if (string.IsNullOrEmpty(WR_RECORDING_STATION))
                {
                    bcPrecNew.WR_RECORDING_STATION = null;
                }
                else
                {
                    bcPrecNew.WR_RECORDING_STATION = WR_RECORDING_STATION;
                }
                if (bcPrecNew.WR_RECORDING_STATION != null)
                {
                    var SN_ID = (from c in bcClimateSiteList
                                 where c.BCText == WR_RECORDING_STATION
                                 select new { c.BCText, c.ClimateSiteName, c.ClimateID }).FirstOrDefault();

                    if (SN_ID.ClimateSiteName == null)
                    {
                        bcPrecNew.WR_RECORDING_STATION_CL = null;
                    }
                    else
                    {
                        bcPrecNew.WR_RECORDING_STATION_CL = SN_ID.ClimateSiteName;
                    }
                    if (SN_ID.ClimateID == null)
                    {
                        bcPrecNew.ClimateID = null;
                    }
                    else
                    {
                        bcPrecNew.ClimateID = SN_ID.ClimateID;
                    }
                }
                bcPrecList.Add(bcPrecNew);
            }

            reader.Close();
            conn.Close();

            int InDBAlready = 0;
            int skip = 0;

            lblTotal.Text = (bcPrecList.Count() + InDBAlready).ToString();
            int take = bcPrecList.Count() / DoBWCount;
            for (int i = 0; i < DoBWCount; i++)
            {
                BWBCPrec bwBCPrec = new BWBCPrec();
                bwBCPrec.bcPrecList = bcPrecList;
                bwBCPrec.skip = skip;
                if (i == (DoBWCount - 1))
                {
                    take = 1000000;
                }
                bwBCPrec.take = take;
                bwListBCPrec[i].RunWorkerAsync(bwBCPrec);
                skip += take;
            }

        }
        private void LoadBCWeatherMissing()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                int countBCPrec = (from c in dbDT.BCPrecipitations
                                   select c).Count();

                if (countBCPrec == 0)
                {
                    richTextBox1.AppendText("LoadBCWeather_Readings has to be run first.");
                    return;
                }
                Application.DoEvents();

                lblTotal.Text = (from c in dbDT.BCPrecipitations
                                 where c.ClimateID == null
                                 select c).Count().ToString();

                Application.DoEvents();

                List<BCPrecipitation> bcPrecList = (from c in dbDT.BCPrecipitations
                                                    where c.ClimateID == null
                                                    select c).ToList<BCPrecipitation>();

                foreach (BCPrecipitation bcp in bcPrecList)
                {
                    Application.DoEvents();

                    string sectorName = (from s in dbDT.BCSurveys
                                         where s.S_ID_NUMBER == bcp.WR_SURVEY
                                         select s.S_SHELLFISH_SECTOR).FirstOrDefault<string>();

                    if (sectorName == "VI04")
                    {
                        bcp.WR_RECORDING_STATION_CL = "VICTORIA INT'L A";
                        bcp.ClimateID = "1018620";
                    }
                    else
                    {
                        BCPrecipitation bcPrecipitationWithClimateID = (from p in dbDT.BCPrecipitations
                                                                        from s in dbDT.BCSurveys
                                                                        where p.WR_SURVEY == s.S_ID_NUMBER
                                                                        && s.S_SHELLFISH_SECTOR == sectorName
                                                                        && p.ClimateID != null
                                                                        select p).FirstOrDefault<BCPrecipitation>();

                        if (bcPrecipitationWithClimateID == null)
                        {
                            richTextBox1.AppendText("Could not find a MarineSamples which has a valid ClimateID. WR_SURVEY is [" + bcp.WR_SURVEY + "]\r\n");
                            continue;
                        }

                        bcp.WR_RECORDING_STATION_CL = bcPrecipitationWithClimateID.WR_RECORDING_STATION_CL;
                        bcp.ClimateID = bcPrecipitationWithClimateID.ClimateID;
                    }
                    dbDT.SaveChanges();

                }

            }
        }
        private void LoadNomenclature()
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                // Delete all Nomenclature in SQL
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000) FROM Nomenclature");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('Nomenclature', RESEED, 0)");

                string connStrNomenclature = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=C:\DataTool\ImportDB\Locator Nomenclature.xls;Extended Properties=Excel 8.0;";

                Application.DoEvents();

                OleDbCommand cmdNomenclature = new OleDbCommand("Select * from [Nomenclature$] order by [LOCATOR];");
                OleDbConnection connNomenclature = new OleDbConnection(connStrNomenclature);

                connNomenclature.Open();
                OleDbDataReader reader;
                cmdNomenclature.Connection = connNomenclature;
                reader = cmdNomenclature.ExecuteReader();
                while (reader.Read())
                {
                    Application.DoEvents();

                    string Locator = "";
                    string Area = "";
                    string Sector = "";
                    string DFO = "";
                    string Subsector = "";
                    string SubsectorDesc = "";
                    string Map = "";
                    DateTime? Updated = new DateTime(2050, 1, 1);
                    if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                    {
                        Locator = "";
                    }
                    else
                    {
                        Locator = reader.GetValue(0).ToString().Trim();
                    }
                    if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                    {
                        Area = "";
                    }
                    else
                    {
                        Area = reader.GetValue(1).ToString().Trim();
                    }
                    if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                    {
                        Sector = "";
                    }
                    else
                    {
                        Sector = reader.GetValue(2).ToString().Trim();
                    }
                    if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                    {
                        DFO = "";
                    }
                    else
                    {
                        DFO = reader.GetValue(3).ToString().Trim();
                    }
                    if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                    {
                        Subsector = "";
                    }
                    else
                    {
                        Subsector = reader.GetValue(4).ToString().Trim();
                    }
                    if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                    {
                        SubsectorDesc = "";
                    }
                    else
                    {
                        SubsectorDesc = reader.GetValue(5).ToString().Trim();
                    }
                    if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                    {
                        Map = "";
                    }
                    else
                    {
                        Map = reader.GetValue(6).ToString().Trim();
                    }
                    if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                    {
                        Updated = null;
                    }
                    else
                    {
                        Updated = DateTime.Parse(reader.GetValue(7).ToString().Trim());
                    }

                    Nomenclature nomenclatureNew = new Nomenclature();
                    nomenclatureNew.Locator = Locator.Trim();
                    nomenclatureNew.Area = Area.Trim();
                    nomenclatureNew.Sector = Sector.Trim();
                    nomenclatureNew.DFO = DFO.Trim();
                    nomenclatureNew.Subsector = Subsector.Trim();
                    nomenclatureNew.SubsectorDesc = SubsectorDesc.Trim();
                    nomenclatureNew.Map = Map.Trim();
                    nomenclatureNew.Updated = Updated;

                    dbDT.Nomenclatures.Add(nomenclatureNew);
                    dbDT.SaveChanges();
                    richTextBox1.AppendText(Locator + "\r\n");
                }
            }
        }
        private void LoadSanitaryDon(string Prov)
        {
            OleDbConnection co = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\CSSP\sanitary.mdb");
            co.Open();

            //List<string> ProvList = new List<string>() { "PEI" };

            List<SanitaryDonOb> obList = new List<SanitaryDonOb>();
            List<SanitaryDonSite> siteList = new List<SanitaryDonSite>();

            string Select = "SELECT * FROM [" + Prov + " sites]";
            OleDbCommand cmd = new OleDbCommand(Select, co);
            OleDbDataReader reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            FieldNameList = new List<string>() { "siteid", "Locator", "Site", "Zone", "easting", "northing", "Datum", "latitude", "longitude" };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(Prov + " site " + reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();

            OleDbCommand cmd2 = new OleDbCommand("SELECT * FROM [" + Prov + " obs]", co);
            OleDbDataReader reader2 = cmd2.ExecuteReader();

            reader2.Read();
            List<string> FieldNameList2 = new List<string>();
            if (Prov == "NB")
            {
                FieldNameList2 = new List<string>() { "id", "Active", "siteid", "Date_Site", "Name_Inspector", "Type", "Status", "Risk_Assessment", "Observations" };
            }
            else
            {
                FieldNameList2 = new List<string>() { "id", "siteid", "Active", "Name_Inspector", "Date_Site", "Type", "Status", "Risk_Assessment", "Observations" };
            }
            for (int i = 0; i < reader2.FieldCount; i++)
            {
                if (reader2.GetName(i) != FieldNameList2[i])
                {
                    richTextBox1.AppendText(Prov + " obs " + reader2.GetName(i) + " is not equal to " + FieldNameList2[i] + "\r\n");
                    return;
                }
            }
            reader2.Close();



            cmd = new OleDbCommand("SELECT * FROM [" + Prov + " obs]", co);
            reader = cmd.ExecuteReader();

            int id = 0;
            bool Active = true;
            int siteid = -1;
            DateTime Date_Site = new DateTime(2050, 1, 1);
            string Name_Inspector = "";
            string Type = "";
            string Status = "";
            string Risk_Assessment = "";
            string Observations = "";

            int Count = 0;
            while (reader.Read())
            {
                Count += 1;
                lblStatus.Text = Prov + " Reading obs " + Count.ToString();
                lblStatus.Refresh();
                Application.DoEvents();

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    id = 0;
                }
                else
                {
                    id = (int)(reader.GetValue(0));
                }

                if (Prov == "NB")
                {
                    if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                    {
                        Active = false;
                    }
                    else
                    {
                        Active = (bool)(reader.GetValue(1));
                    }

                    if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                    {
                        siteid = -1;
                    }
                    else
                    {
                        siteid = (int)(reader.GetValue(2));
                    }
                    if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                    {
                        Date_Site = new DateTime(2050, 1, 1);
                    }
                    else
                    {
                        Date_Site = (DateTime)(reader.GetValue(3));
                    }

                    if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                    {
                        Name_Inspector = "";
                    }
                    else
                    {
                        Name_Inspector = reader.GetValue(4).ToString().Trim();
                    }
                }
                else
                {
                    if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                    {
                        Active = false;
                    }
                    else
                    {
                        Active = (bool)(reader.GetValue(2));
                    }

                    if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                    {
                        siteid = -1;
                    }
                    else
                    {
                        siteid = (int)(reader.GetValue(1));
                    }
                    if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                    {
                        Date_Site = new DateTime(2050, 1, 1);
                    }
                    else
                    {
                        Date_Site = (DateTime)(reader.GetValue(4));
                    }

                    if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                    {
                        Name_Inspector = "";
                    }
                    else
                    {
                        Name_Inspector = reader.GetValue(3).ToString().Trim();
                    }
                }


                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    Type = "";
                }
                else
                {
                    Type = reader.GetValue(5).ToString().Trim();
                }

                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    Status = "";
                }
                else
                {
                    Status = reader.GetValue(6).ToString().Trim();
                }

                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    Risk_Assessment = "";
                }
                else
                {
                    Risk_Assessment = reader.GetValue(7).ToString().Trim();
                }

                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    Observations = "";
                }
                else
                {
                    Observations = reader.GetValue(8).ToString().Trim();
                }

                SanitaryDonOb obNew = new SanitaryDonOb();
                obNew.Active = Active;
                obNew.siteid = siteid;
                obNew.ObsDate = Date_Site;
                obNew.Name_Inspector = Name_Inspector;
                obNew.Type = Type;
                obNew.Status = Status;
                obNew.Risk_Assessment = Risk_Assessment;
                obNew.Observations = Observations;

                obList.Add(obNew);
            }
            reader.Close();

            Select = "SELECT * FROM [" + Prov + " sites]";
            cmd = new OleDbCommand(Select, co);
            reader = cmd.ExecuteReader();

            Count = 0;
            while (reader.Read())
            {
                Count += 1;
                lblStatus.Text = Prov + " Reading sites " + Count.ToString();
                lblStatus.Refresh();
                Application.DoEvents();

                siteid = -1;
                string Locator = "";
                int Site = -1;
                int Zone = -1;
                float easting = -1f;
                float northing = -1f;
                string Datum = "";
                float latitude = -1f;
                float longitude = -1f;

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    siteid = -1;
                }
                else
                {
                    siteid = (int)(reader.GetValue(0));
                }
                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    Locator = "";
                }
                else
                {
                    Locator = reader.GetValue(1).ToString().Trim();
                }
                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    Site = -1;
                }
                else
                {
                    Site = (int)(double)(reader.GetValue(2));
                }
                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    Zone = -1;
                }
                else
                {
                    Zone = (int)(reader.GetValue(3));
                }
                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    easting = -1;
                }
                else
                {
                    easting = (float)(double)(reader.GetValue(4));
                }
                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    northing = -1;
                }
                else
                {
                    northing = (float)(double)(reader.GetValue(5));
                }
                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    Datum = "";
                }
                else
                {
                    Datum = reader.GetValue(6).ToString().Trim();
                }
                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    latitude = -1;
                }
                else
                {
                    latitude = (float)(double)(reader.GetValue(7));
                }
                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    longitude = -1;
                }
                else
                {
                    longitude = (float)(double)(reader.GetValue(8));
                }

                SanitaryDonSite siteNew = new SanitaryDonSite();
                siteNew.siteid = siteid;
                siteNew.Locator = Locator;
                siteNew.Site = Site;
                siteNew.Zone = Zone;
                siteNew.easting = easting;
                siteNew.northing = northing;
                siteNew.Datum = Datum;
                siteNew.latitude = latitude;
                siteNew.longitude = longitude;

                foreach (SanitaryDonOb ob in obList.Where(c => c.siteid == siteid))
                {
                    siteNew.SanitaryDonObs.Add(ob);
                }

                siteList.Add(siteNew);
            }
            reader.Close();

            int InDBAlready = 0;

            using (TempDataToolDBEntities dbTD = new TempDataToolDBEntities())
            {
                InDBAlready = (from c in dbTD.SanitaryDonSites
                               select c).Count();
            }

            int skip = 0;

            lblTotal.Text = (siteList.Count() + InDBAlready).ToString();
            int take = siteList.Count() / DoBWCount;
            for (int i = 0; i < DoBWCount; i++)
            {
                BWDon bwDon = new BWDon();
                bwDon.siteList = siteList;
                bwDon.skip = skip;
                if (i == (DoBWCount - 1))
                {
                    take = 1000000;
                }
                bwDon.take = take;
                bwListDon[i].RunWorkerAsync(bwDon);
                skip += take;
            }
        }
        private void LoadSanitaryPatrice()
        {
            string StartDirName = @"C:\Sanitary_Excel_Patrice\";

            DirectoryInfo di = new DirectoryInfo(StartDirName);

            FileInfo[] files = di.GetFiles();
            int TotalFileCount = files.Count();
            for (int i = 0; i < TotalFileCount; i++)
            {
                BWPat bwPat = new BWPat();
                bwPat.FileName = files[i].FullName;
                richTextBox1.AppendText("Doing [" + files[i].FullName + "]\r\n");

                //DoPatriceFunction(bwPat);
                bwListPat[i].RunWorkerAsync(bwPat);
            }
        }
        private void LoadSanitaryDav(string Prov)
        {
            OleDbConnection co = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\CSSP\sanitary_Ryan.mdb");
            co.Open();

            List<SanitaryDavOb> obList = new List<SanitaryDavOb>();
            List<SanitaryDavSite> siteList = new List<SanitaryDavSite>();

            string Select = "SELECT * FROM [" + Prov + "]"; // Prov is NS or PEI
            OleDbCommand cmd = new OleDbCommand(Select, co);
            OleDbDataReader reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            if (Prov == "NS")
            {
                FieldNameList = new List<string>() { "ID", "Locator", "Locator_Name", "Locator_LocatorName", "Site", "Type", "Observations",
                "Zone", "Datum", "latitude", "longitude", "Active", "Date_Created", "Date_Updated", "Photo", "Civic", "Sample_MPN_per_100ml", "Additional_Info" };
            }
            else
            {
                FieldNameList = new List<string>() { "ID", "Locator", "Locator_Name", "Locator_LocatorName", "Site", "Type", "Observations",
                "Active", "NAD", "Zone", "latitude_NAD83", "longitude_NAD83", "Date_Created", "Date_Updated", "Photo", "Civic", "Sample_MPN_per_100mL", "Additional_Info" };
            }
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(Prov + " " + reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();

            reader = cmd.ExecuteReader();

            cmd = new OleDbCommand("SELECT * FROM [" + Prov + "]", co);
            reader = cmd.ExecuteReader();

            int ID = -1;                                       //0
            string Locator = "";                               //1
            string Locator_Name = "";                          //2
            string Locator_LocatorName = "";                   //3
            int Site = -1;                                     //4
            string Type = "";                                  //5
            string Observations = "";                          //6
            int Zone = -1;                                     //7
            string Datum = "";                                 //8
            float latitude = 0.0f;                             //9
            float longitude = 0.0f;                            //10
            bool Active = true;                                //11
            DateTime Date = new DateTime(2050, 1, 1);          //12
            DateTime Updated = new DateTime(2050, 1, 1);       //13
            //string Photo = "";                                 //14
            //string Civic = "";                                 //15
            //string Sample_MPN_Per_100ml = "";                  //16
            //string Detail = "";                                //17

            int Count = 0;
            while (reader.Read())
            {
                Count += 1;

                lblStatus.Text = Prov + " Reading obs " + Count.ToString();
                lblStatus.Refresh();
                Application.DoEvents();

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    ID = 0;
                }
                else
                {
                    ID = (int)(reader.GetValue(0));
                }
                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    Locator = "";
                }
                else
                {
                    Locator = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    Locator_Name = "";
                }
                else
                {
                    Locator_Name = reader.GetValue(2).ToString().Trim();
                }
                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    Locator_LocatorName = "";
                }
                else
                {
                    Locator_LocatorName = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    Site = -1;
                }
                else
                {
                    Site = int.Parse(reader.GetValue(4).ToString().Trim());
                }
                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    Type = "";
                }
                else
                {
                    Type = reader.GetValue(5).ToString().Trim();
                }
                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    Observations = "";
                }
                else
                {
                    Observations = reader.GetValue(6).ToString().Trim();
                }

                if (Prov == "NS")
                {
                    if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                    {
                        Zone = -1;
                    }
                    else
                    {
                        Zone = int.Parse(reader.GetValue(7).ToString().Trim());
                    }
                    if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                    {
                        Datum = "";
                    }
                    else
                    {
                        Datum = reader.GetValue(8).ToString().Trim();
                    }
                    if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                    {
                        latitude = -1f;
                    }
                    else
                    {
                        latitude = float.Parse(reader.GetValue(9).ToString());
                    }
                    if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                    {
                        longitude = -1f;
                    }
                    else
                    {
                        longitude = float.Parse(reader.GetValue(10).ToString());
                    }
                    if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                    {
                        Active = false;
                    }
                    else
                    {
                        Active = (bool)(reader.GetValue(11));
                    }
                    if (reader.GetValue(12).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(12).ToString()))
                    {
                        Date = new DateTime(2050, 1, 1);
                    }
                    else
                    {
                        Date = (DateTime)(reader.GetValue(12));
                    }
                    if (reader.GetValue(13).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(13).ToString()))
                    {
                        Updated = new DateTime(2050, 1, 1);
                    }
                    else
                    {
                        Updated = (DateTime)(reader.GetValue(13));
                    }
                }
                else
                {
                    if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                    {
                        Active = false;
                    }
                    else
                    {
                        Active = (bool)(reader.GetValue(7));
                    }

                    if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                    {
                        Datum = "";
                    }
                    else
                    {
                        Datum = reader.GetValue(8).ToString().Trim();
                    }
                    if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                    {
                        Zone = -1;
                    }
                    else
                    {
                        Zone = (int)(double)(reader.GetValue(9));
                    }
                    if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                    {
                        latitude = -1f;
                    }
                    else
                    {
                        latitude = float.Parse(reader.GetValue(10).ToString().Trim());
                    }
                    if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                    {
                        longitude = -1f;
                    }
                    else
                    {
                        longitude = float.Parse(reader.GetValue(11).ToString().Trim());
                    }
                    if (reader.GetValue(12).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(12).ToString()))
                    {
                        Date = new DateTime(2050, 1, 1);
                    }
                    else
                    {
                        Date = (DateTime)(reader.GetValue(12));
                    }
                    if (reader.GetValue(13).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(13).ToString()))
                    {
                        Updated = new DateTime(2050, 1, 1);
                    }
                    else
                    {
                        Updated = (DateTime)(reader.GetValue(13));
                    }
                }

                SanitaryDavSite sanitaryDavSiteExist = siteList.Where(c => c.Locator == Locator && c.Site == Site).FirstOrDefault<SanitaryDavSite>();
                if (sanitaryDavSiteExist == null)
                {
                    SanitaryDavSite siteNew = new SanitaryDavSite();
                    siteNew.siteid = null;
                    siteNew.Locator = Locator;
                    siteNew.Site = Site;
                    siteNew.Zone = Zone;
                    siteNew.easting = null;
                    siteNew.northing = null;
                    siteNew.Datum = Datum;
                    siteNew.latitude = latitude;
                    siteNew.longitude = longitude;

                    siteList.Add(siteNew);

                    SanitaryDavOb obNew = new SanitaryDavOb();
                    obNew.Active = Active;
                    obNew.siteid = null;
                    if (Updated != null)
                    {
                        obNew.ObsDate = Updated;
                    }
                    else
                    {
                        obNew.ObsDate = Date;
                    }
                    obNew.Name_Inspector = null;
                    obNew.Type = Type;
                    obNew.Status = null;
                    obNew.Risk_Assessment = null;
                    obNew.Observations = Observations;

                    siteNew.SanitaryDavObs.Add(obNew);


                }
                else
                {
                    SanitaryDavOb obNew = new SanitaryDavOb();
                    obNew.Active = Active;
                    obNew.siteid = null;
                    if (Updated != null)
                    {
                        obNew.ObsDate = Updated;
                    }
                    else
                    {
                        obNew.ObsDate = Date;
                    }
                    obNew.Name_Inspector = null;
                    obNew.Type = Type;
                    obNew.Status = null;
                    obNew.Risk_Assessment = null;
                    obNew.Observations = Observations;

                    sanitaryDavSiteExist.SanitaryDavObs.Add(obNew);

                }
            }
            reader.Close();

            int InDBAlready = 0;

            using (TempDataToolDBEntities dbTD = new TempDataToolDBEntities())
            {
                InDBAlready = (from c in dbTD.SanitaryDavSites
                               select c).Count();
            }

            int skip = 0;

            lblTotal.Text = (siteList.Count() + InDBAlready).ToString();
            int take = siteList.Count() / DoBWCount;
            for (int i = 0; i < DoBWCount; i++)
            {
                BWDav bwDav = new BWDav();
                bwDav.siteList = siteList;
                bwDav.skip = skip;
                if (i == (DoBWCount - 1))
                {
                    take = 1000000;
                }
                bwDav.take = take;
                bwListDav[i].RunWorkerAsync(bwDav);
                skip += take;
            }

        }
        private void LoadSanitaryJoe(string Prov)
        {
            OleDbConnection co = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\CSSP\sanitary_Joe.mdb");
            co.Open();

            List<SanitaryJoeOb> obList = new List<SanitaryJoeOb>();
            List<SanitaryJoeSite> siteList = new List<SanitaryJoeSite>();

            string Select = "SELECT * FROM [" + Prov + "]"; // Prov is NB
            OleDbCommand cmd = new OleDbCommand(Select, co);
            OleDbDataReader reader = cmd.ExecuteReader();

            reader.Read();

            List<string> FieldNameList = new List<string>();
            if (Prov == "NB")
            {
                FieldNameList = new List<string>() { "ID", "Locator", "Locator_Name", "Locator_LocatorName", "Site", "Type", "Observations",
                 "Zone", "Datum", "latitude", "longitude", "Active", "Date_Created", "Date_Updated", "Photo", "Civic", "Sample_MPN_Per_100ml", "Additional_Info" };
            }
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetName(i) != FieldNameList[i])
                {
                    richTextBox1.AppendText(Prov + " " + reader.GetName(i) + " is not equal to " + FieldNameList[i] + "\r\n");
                    return;
                }
            }
            reader.Close();

            reader = cmd.ExecuteReader();

            cmd = new OleDbCommand("SELECT * FROM [" + Prov + "]", co);
            reader = cmd.ExecuteReader();

            int ID = -1;                                       //0
            string Locator = "";                               //1
            string Locator_Name = "";                          //2
            string Locator_LocatorName = "";                   //3
            int Site = -1;                                     //4
            string Type = "";                                  //5
            string Observations = "";                          //6
            int Zone = -1;                                     //7
            string Datum = "";                                 //8
            float latitude = 0.0f;                             //9
            float longitude = 0.0f;                            //10
            bool Active = true;                                //11
            DateTime Date = new DateTime(2050, 1, 1);          //12
            DateTime Updated = new DateTime(2050, 1, 1);       //13
            //string Photo = "";                                 //14
            //string Civic = "";                                 //15
            //string Sample_MPN_Per_100ml = "";                  //16
            //string Detail = "";                                //17

            int Count = 0;
            while (reader.Read())
            {
                Count += 1;

                lblStatus.Text = Prov + " Reading obs " + Count.ToString();
                lblStatus.Refresh();
                Application.DoEvents();

                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    ID = 0;
                }
                else
                {
                    ID = (int)(reader.GetValue(0));
                }
                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    Locator = "";
                }
                else
                {
                    Locator = reader.GetValue(1).ToString().Trim();
                }

                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    Locator_Name = "";
                }
                else
                {
                    Locator_Name = reader.GetValue(2).ToString().Trim();
                }
                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    Locator_LocatorName = "";
                }
                else
                {
                    Locator_LocatorName = reader.GetValue(3).ToString().Trim();
                }

                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    Site = -1;
                }
                else
                {
                    Site = int.Parse(reader.GetValue(4).ToString().Trim());
                }
                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    Type = "";
                }
                else
                {
                    Type = reader.GetValue(5).ToString().Trim();
                }
                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    Observations = "";
                }
                else
                {
                    Observations = reader.GetValue(6).ToString().Trim();
                }
                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    Zone = -1;
                }
                else
                {
                    Zone = int.Parse(reader.GetValue(7).ToString().Trim());
                }
                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    Datum = "";
                }
                else
                {
                    Datum = reader.GetValue(8).ToString().Trim();
                }
                if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                {
                    latitude = -1f;
                }
                else
                {
                    latitude = float.Parse(reader.GetValue(9).ToString());
                }
                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                {
                    longitude = -1f;
                }
                else
                {
                    longitude = float.Parse(reader.GetValue(10).ToString());
                }
                if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                {
                    Active = false;
                }
                else
                {
                    Active = (bool)(reader.GetValue(11));
                }
                if (reader.GetValue(12).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(12).ToString()))
                {
                    Date = new DateTime(2050, 1, 1);
                }
                else
                {
                    Date = (DateTime)(reader.GetValue(12));
                }
                if (reader.GetValue(13).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(13).ToString()))
                {
                    Updated = new DateTime(2050, 1, 1);
                }
                else
                {
                    Updated = (DateTime)(reader.GetValue(13));
                }



                SanitaryJoeSite sanitaryJoeSiteExist = siteList.Where(c => c.Locator == Locator && c.Site == Site).FirstOrDefault<SanitaryJoeSite>();
                if (sanitaryJoeSiteExist == null)
                {
                    SanitaryJoeSite siteNew = new SanitaryJoeSite();
                    siteNew.siteid = null;
                    siteNew.Locator = Locator;
                    siteNew.Site = Site;
                    siteNew.Zone = Zone;
                    siteNew.easting = null;
                    siteNew.northing = null;
                    siteNew.Datum = Datum;
                    siteNew.latitude = latitude;
                    siteNew.longitude = longitude;

                    siteList.Add(siteNew);

                    SanitaryJoeOb obNew = new SanitaryJoeOb();
                    obNew.Active = Active;
                    obNew.siteid = null;
                    if (Updated != null)
                    {
                        obNew.ObsDate = Updated;
                    }
                    else
                    {
                        obNew.ObsDate = Date;
                    }
                    obNew.Name_Inspector = null;
                    obNew.Type = Type;
                    obNew.Status = null;
                    obNew.Risk_Assessment = null;
                    obNew.Observations = Observations;

                    siteNew.SanitaryJoeObs.Add(obNew);


                }
                else
                {
                    SanitaryJoeOb obNew = new SanitaryJoeOb();
                    obNew.Active = Active;
                    obNew.siteid = null;
                    if (Updated != null)
                    {
                        obNew.ObsDate = Updated;
                    }
                    else
                    {
                        obNew.ObsDate = Date;
                    }
                    obNew.Name_Inspector = null;
                    obNew.Type = Type;
                    obNew.Status = null;
                    obNew.Risk_Assessment = null;
                    obNew.Observations = Observations;

                    sanitaryJoeSiteExist.SanitaryJoeObs.Add(obNew);

                }
            }
            reader.Close();

            int InDBAlready = 0;

            using (TempDataToolDBEntities dbTD = new TempDataToolDBEntities())
            {
                InDBAlready = (from c in dbTD.SanitaryJoeSites
                               select c).Count();
            }

            int skip = 0;

            lblTotal.Text = (siteList.Count() + InDBAlready).ToString();
            int take = siteList.Count() / DoBWCount;
            for (int i = 0; i < DoBWCount; i++)
            {
                BWJoe bwJoe = new BWJoe();
                bwJoe.siteList = siteList;
                bwJoe.skip = skip;
                if (i == (DoBWCount - 1))
                {
                    take = 1000000;
                }
                bwJoe.take = take;
                bwListJoe[i].RunWorkerAsync(bwJoe);
                skip += take;
            }

        }
        private void LoadSanitaryDaveC()
        {

            FileInfo fi = new FileInfo(@"C:\CSSP\CSSPWebTools_docs_and_templates\San Data NL 02-04-2015.xlsx");

            if (!fi.Exists)
            {
                richTextBox1.AppendText("Error Could not find file [" + fi.FullName + "]");
            }

            BWDavC bwDavC = new BWDavC();
            bwDavC.FileName = fi.FullName;

            bwListDavC[0].RunWorkerAsync(bwDavC);
        }
        #endregion Functions

        #region Classes
        private class BCClimateSite
        {
            public string BCText { get; set; }
            public string ClimateSiteName { get; set; }
            public string ClimateID { get; set; }
        }
        private class BWBCRun
        {
            public List<BCSurvey> bcRunList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWBCPrec
        {
            public List<BCPrecipitation> bcPrecList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWBCMarineStation
        {
            public List<BCMarineSampleStation> bcMarineStationList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWBCMarineSample
        {
            public List<BCMarineSample> bcMarineSampleList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWBCLandStation
        {
            public List<BCLandSampleStation> bcLandStationList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWBCLandSample
        {
            public List<BCLandSample> bcLandSampleList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWBCPolSource
        {
            public List<BCPolSource> bcPolSourceList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWDon
        {
            public List<SanitaryDonSite> siteList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWDav
        {
            public List<SanitaryDavSite> siteList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWJoe
        {
            public List<SanitaryJoeSite> siteList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWDavC
        {
            public string FileName { get; set; }
        }
        private class BWPat
        {
            public string FileName { get; set; }
        }
        private class BWRun
        {
            public List<ASGADRun> asgadRunList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWPrecData
        {
            public List<ASGADPrecData> asgadPrecDataList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWSample
        {
            public List<ASGADSample> asgadSampleList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWStation
        {
            public List<ASGADStation> asgadStationList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class BWSubsect
        {
            public List<ASGADSubsect> asgadSubsectList { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        private class StnInfo
        {
            public string Prov { get; set; }
            public string Area { get; set; }
            public string Sector { get; set; }
            public string Subsector { get; set; }
            public string Stn_Nbr { get; set; }
        }
        private class RunInfo
        {
            public string Prov { get; set; }
            public string Area { get; set; }
            public string Sector { get; set; }
            public string Subsector { get; set; }
            public int? Run_Nbr { get; set; }
        }
        #endregion Classes

        private void button1_Click(object sender, EventArgs e)
        {
            List<LocationNameLatLng> locationNameLatLng = new List<LocationNameLatLng>();

            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                locationNameLatLng = (from c in dbDT.LocationNameLatLngs
                                      where c.Municipality != ""
                                      select c).ToList<LocationNameLatLng>();

                foreach (var v in locationNameLatLng)
                {
                    if (v.Lat == -1)
                    {
                        richTextBox1.AppendText("-----------1\t" + v.Municipality + "\r\n");
                    }
                }

            }


            List<string> provList = new List<string>() { "British Columbia", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec" };
            foreach (string prov in provList)
            {
                using (CSSPAppDB2Entities dbMVC = new CSSPAppDB2Entities())
                {
                    var csspItemAndLanguageList = (from c in dbMVC.CSSPItems
                                                   from cl in dbMVC.CSSPItemLanguages
                                                   from cp in dbMVC.CSSPItemPaths
                                                   where c.CSSPItemID == cl.CSSPItemID
                                                   && c.CSSPItemID == cp.CSSPItemID
                                                   && cl.Language == "en"
                                                   && cl.CSSPItemText == prov
                                                   select new { c, cl, cp }).ToList();

                    foreach (var v in csspItemAndLanguageList)
                    {
                        int level = v.cp.Level + 2;
                        richTextBox1.AppendText(v.cl.CSSPItemText + "\r\n\r\n");
                        var csspItemAndLangListMuni = (from c in dbMVC.CSSPItems
                                                       from cl in dbMVC.CSSPItemLanguages
                                                       from cp in dbMVC.CSSPItemPaths
                                                       where c.CSSPItemID == cl.CSSPItemID
                                                       && c.CSSPItemID == cp.CSSPItemID
                                                       && cl.Language == "en"
                                                       && cp.Path.StartsWith(v.cp.Path)
                                                       && cp.Level == level
                                                       select new { c, cl }).ToList();

                        foreach (var vv in csspItemAndLangListMuni)
                        {
                            if (locationNameLatLng.Where(c => c.Province == prov && c.Municipality == vv.cl.CSSPItemText).Any())
                            {
                                richTextBox1.AppendText("EXIST: " + vv.cl.CSSPItemText + "\r\n");
                            }
                            else
                            {
                                richTextBox1.AppendText("------- no EXIST: " + vv.cl.CSSPItemText + "\r\n");

                                using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
                                {
                                    //LocationNameLatLng locNew = new LocationNameLatLng();
                                    //locNew.Municipality = vv.cl.CSSPItemText.Trim();
                                    //locNew.Province = v.cl.CSSPItemText.Trim();
                                    //locNew.Lat = -1;
                                    //locNew.Lng = -1;

                                    //dbDT.LocationNameLatLngs.Add(locNew);

                                    //dbDT.SaveChanges();
                                }
                            }
                        }
                    }

                }
            }
        }

        private void butLoadQCClimateSite_Click(object sender, EventArgs e)
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM QCClimateSites");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('QCClimateSites', RESEED, 0)");
            }

            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                List<SanitaryPatSite> sitePatList = new List<SanitaryPatSite>();

                string FileName = @"C:\DataTool\ImportDB\QC Climate Sites.xlsx";
                string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties=Excel 12.0";
                OleDbConnection conn = new OleDbConnection(connectionString);

                conn.Open();
                OleDbDataReader reader;

                OleDbCommand comm = new OleDbCommand("Select * from [Sheet1$];");

                try
                {
                    comm.Connection = conn;
                    reader = comm.ExecuteReader();
                }
                catch (Exception) // not a real error just trying to find the year of the sheet
                {
                    richTextBox1.AppendText("Could not read QC Climate Sites.xlsx");
                    return;
                }

                List<string> FieldNameList = new List<string>();
                FieldNameList = new List<string>() { "QCMeteo", "ClimateName", "ClimateID" };
                for (int j = 0; j < reader.FieldCount; j++)
                {
                    if (reader.GetName(j) != FieldNameList[j])
                    {
                        richTextBox1.AppendText(FileName + " Sheet1 " + reader.GetName(j) + " is not equal to " + FieldNameList[j] + "\r\n");
                        return;
                    }
                }
                reader.Close();

                reader = comm.ExecuteReader();


                while (reader.Read())
                {
                    string QCMeteo = "";
                    string ClimateName = "";
                    string ClimateID = "";

                    if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                    {
                        QCMeteo = "";
                        richTextBox1.AppendText("Could not read QCMeteo");
                        return;
                    }
                    else
                    {
                        QCMeteo = reader.GetValue(0).ToString().Trim();
                    }
                    if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                    {
                        ClimateName = "";
                        richTextBox1.AppendText("Could not read QCMeteo");
                        return;
                    }
                    else
                    {
                        ClimateName = reader.GetValue(1).ToString().Trim();
                    }
                    if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                    {
                        ClimateID = "";
                        richTextBox1.AppendText("Could not read QCMeteo");
                        return;
                    }
                    else
                    {
                        ClimateID = reader.GetValue(2).ToString().Trim();
                    }

                    richTextBox1.AppendText(QCMeteo + "\t" + ClimateName + "\t" + ClimateID + "\r\n");

                    QCClimateSite qcClimateSiteNew = new QCClimateSite();
                    qcClimateSiteNew.station_meteo = QCMeteo;
                    qcClimateSiteNew.ClimateSiteName = ClimateName;
                    qcClimateSiteNew.ClimateID = ClimateID;

                    dbDT.QCClimateSites.Add(qcClimateSiteNew);

                    dbDT.SaveChanges();
                }
            }
        }

        private void butBCAreaSectorSubsectorNames_Click(object sender, EventArgs e)
        {
            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                dbDT.Database.ExecuteSqlCommand("DELETE TOP (10000000) FROM BCNames");
                dbDT.Database.ExecuteSqlCommand("DBCC CHECKIDENT('BCNames', RESEED, 0)");
            }

            using (TempDataToolDBEntities dbDT = new TempDataToolDBEntities())
            {
                List<SanitaryPatSite> sitePatList = new List<SanitaryPatSite>();

                string FileName = @"C:\DataTool\ImportDB\BC Zone_GrowingArea_Sector_List_20140422.xlsx";
                string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties=Excel 12.0";
                OleDbConnection conn = new OleDbConnection(connectionString);

                conn.Open();
                OleDbDataReader reader;

                OleDbCommand comm = new OleDbCommand("Select * from [Sheet1$];");

                try
                {
                    comm.Connection = conn;
                    reader = comm.ExecuteReader();
                }
                catch (Exception) // not a real error just trying to find the year of the sheet
                {
                    richTextBox1.AppendText("Could not read QC Climate Sites.xlsx");
                    return;
                }

                List<string> FieldNameList = new List<string>();
                FieldNameList = new List<string>() { "Zone", "ZoneName", "GrowingArea", "GrowingAreaName", "Sector", "SectorName" };
                for (int j = 0; j < reader.FieldCount; j++)
                {
                    if (reader.GetName(j) != FieldNameList[j])
                    {
                        richTextBox1.AppendText(FileName + " Sheet1 " + reader.GetName(j) + " is not equal to " + FieldNameList[j] + "\r\n");
                        return;
                    }
                }
                reader.Close();

                reader = comm.ExecuteReader();

                string OldZone = "";
                string OldGrowingArea = "";
                string OldSector = "";

                while (reader.Read())
                {
                    string Zone = "";
                    string ZoneName = "";
                    string GrowingArea = "";
                    string GrowingAreaName = "";
                    string Sector = "";
                    string SectorName = "";

                    if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                    {
                        Zone = "";
                        richTextBox1.AppendText("Could not read Zone");
                        return;
                    }
                    else
                    {
                        Zone = reader.GetValue(0).ToString().Trim();
                    }
                    if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                    {
                        ZoneName = "";
                        richTextBox1.AppendText("Could not read ZoneName");
                        return;
                    }
                    else
                    {
                        ZoneName = reader.GetValue(1).ToString().Trim();
                    }
                    if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                    {
                        GrowingArea = "";
                        richTextBox1.AppendText("Could not read GrowingArea");
                        return;
                    }
                    else
                    {
                        GrowingArea = reader.GetValue(2).ToString().Trim();
                    }
                    if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                    {
                        GrowingAreaName = "";
                        richTextBox1.AppendText("Could not read GrowingAreaName");
                        return;
                    }
                    else
                    {
                        GrowingAreaName = reader.GetValue(3).ToString().Trim();
                    }
                    if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                    {
                        Sector = "";
                        richTextBox1.AppendText("Could not read Sector");
                        return;
                    }
                    else
                    {
                        Sector = reader.GetValue(4).ToString().Trim();
                    }
                    if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                    {
                        SectorName = "";
                        richTextBox1.AppendText("Could not read SectorName");
                        return;
                    }
                    else
                    {
                        SectorName = reader.GetValue(5).ToString().Trim();
                    }
                    richTextBox1.AppendText(Zone + "\t" + ZoneName + "\t" + GrowingArea + "\t" + GrowingAreaName + "\t" + Sector + "\t" + SectorName + "\r\n");

                    if (OldZone != Zone)
                    {
                        string Accro = Zone.Replace("Zone1 (", "").Replace("Zone2 (", "").Replace("Zone3 (", "").Replace("Zone4 (", "").Replace(")", "").Trim();
                        string TheName = ZoneName;
                        BCName bcNameNew = new BCName();
                        bcNameNew.Accronym = Accro;
                        bcNameNew.Name = TheName;

                        dbDT.BCNames.Add(bcNameNew);

                        dbDT.SaveChanges();

                        OldZone = Zone;
                    }

                    if (OldGrowingArea != GrowingArea)
                    {
                        string Accro = GrowingArea.Replace("Zone1 (", "").Replace("Zone2 (", "").Replace("Zone3 (", "").Replace("Zone4 (", "").Replace(")", "").Trim();
                        string TheName = GrowingAreaName;
                        BCName bcNameNew = new BCName();
                        bcNameNew.Accronym = Accro;
                        bcNameNew.Name = TheName;

                        dbDT.BCNames.Add(bcNameNew);

                        dbDT.SaveChanges();

                        OldGrowingArea = GrowingArea;
                    }

                    if (OldSector != Sector)
                    {

                        string Accro = Sector.Replace("Zone1 (", "").Replace("Zone2 (", "").Replace("Zone3 (", "").Replace("Zone4 (", "").Replace(")", "").Trim();
                        string TheName = SectorName;
                        BCName bcNameNew = new BCName();
                        bcNameNew.Accronym = Accro;
                        bcNameNew.Name = TheName;

                        dbDT.BCNames.Add(bcNameNew);

                        dbDT.SaveChanges();
                    }

                }
            }
        }

        private void butAddInspectorNames_Click(object sender, EventArgs e)
        {
            TempDataToolDBEntities dbT = new TempDataToolDBEntities();
            List<SanitaryDonSite> sanitaryDonSiteList = (from c in dbT.SanitaryDonSites
                                                         select c).ToList();

            List<SanitaryDavSite> sanitaryDavSiteList = (from c in dbT.SanitaryDavSites
                                                         select c).ToList();

            List<SanitaryPatSite> sanitaryPatSiteList = (from c in dbT.SanitaryPatSites
                                                         select c).ToList();

            List<SanitaryJoeSite> sanitaryJoeSiteList = (from c in dbT.SanitaryJoeSites
                                                         select c).ToList();

            int count = 0;
            int Total = sanitaryDonSiteList.Count;
            foreach (SanitaryDonSite sanitaryDonSite in sanitaryDonSiteList)
            {
                count += 1;
                lblStatus.Text = count + " of " + Total;
                lblStatus.Refresh();
                Application.DoEvents();

                using (TempDataToolDBEntities dbT2 = new TempDataToolDBEntities())
                {
                    if (sanitaryDonSite.Locator.Substring(0, 2) == "NB")
                    {
                        List<SanitaryDonOb> sanitaryDonObList = (from c in dbT.SanitaryDonObs
                                                                 where c.SanitaryDonSiteID == sanitaryDonSite.SanitaryDonSiteID
                                                                 select c).ToList();

                        // doing Pat
                        SanitaryPatSite sanitaryPatSite = (from c in sanitaryPatSiteList
                                                           where c.Locator == sanitaryDonSite.Locator
                                                           && c.Site == sanitaryDonSite.Site
                                                           select c).FirstOrDefault();

                        if (sanitaryPatSite != null)
                        {
                            List<SanitaryPatOb> sanitaryPatObList = (from c in dbT.SanitaryPatObs
                                                                     where c.SanitaryPatSiteID == sanitaryPatSite.SanitaryPatSiteID
                                                                     select c).ToList();

                            var ID_Name_Inspector = (from c in sanitaryDonObList
                                                     from d in sanitaryPatObList
                                                     where c.ObsDate == d.ObsDate
                                                     select new { d.SanitaryPatObsID, c.Name_Inspector }).FirstOrDefault();

                            if (ID_Name_Inspector != null)
                            {
                                SanitaryPatOb sanitaryPatObToChange = (from c in dbT2.SanitaryPatObs
                                                                       where c.SanitaryPatObsID == ID_Name_Inspector.SanitaryPatObsID
                                                                       select c).FirstOrDefault();

                                if (sanitaryPatObToChange != null)
                                {
                                    sanitaryPatObToChange.Name_Inspector = ID_Name_Inspector.Name_Inspector;
                                }
                            }

                        }

                        // doint Joe
                        SanitaryJoeSite sanitaryJoeSite = (from c in sanitaryJoeSiteList
                                                           where c.Locator == sanitaryDonSite.Locator
                                                           && c.Site == sanitaryDonSite.Site
                                                           select c).FirstOrDefault();

                        if (sanitaryJoeSite != null)
                        {
                            List<SanitaryJoeOb> sanitaryJoeObList = (from c in dbT.SanitaryJoeObs
                                                                     where c.SanitaryJoeSiteID == sanitaryJoeSite.SanitaryJoeSiteID
                                                                     select c).ToList();

                            var ID_Name_Inspector = (from c in sanitaryDonObList
                                                     from d in sanitaryJoeObList
                                                     where c.ObsDate == d.ObsDate
                                                     select new { d.SanitaryJoeObsID, c.Name_Inspector }).FirstOrDefault();

                            if (ID_Name_Inspector != null)
                            {
                                SanitaryJoeOb sanitaryJoeObToChange = (from c in dbT2.SanitaryJoeObs
                                                                       where c.SanitaryJoeObsID == ID_Name_Inspector.SanitaryJoeObsID
                                                                       select c).FirstOrDefault();

                                if (sanitaryJoeObToChange != null)
                                {
                                    sanitaryJoeObToChange.Name_Inspector = ID_Name_Inspector.Name_Inspector;
                                }
                            }

                        }

                    }


                    if (sanitaryDonSite.Locator.Substring(0, 2) == "NS")
                    {
                        List<SanitaryDonOb> sanitaryDonObList = (from c in dbT.SanitaryDonObs
                                                                 where c.SanitaryDonSiteID == sanitaryDonSite.SanitaryDonSiteID
                                                                 select c).ToList();

                        // doing Dav
                        SanitaryDavSite sanitaryDavSite = (from c in sanitaryDavSiteList
                                                           where c.Locator == sanitaryDonSite.Locator
                                                           && c.Site == sanitaryDonSite.Site
                                                           select c).FirstOrDefault();

                        if (sanitaryDavSite != null)
                        {
                            List<SanitaryDavOb> sanitaryDavObList = (from c in dbT.SanitaryDavObs
                                                                     where c.SanitaryDavSiteID == sanitaryDavSite.SanitaryDavSiteID
                                                                     select c).ToList();

                            var ID_Name_Inspector = (from c in sanitaryDonObList
                                                     from d in sanitaryDavObList
                                                     where c.ObsDate == d.ObsDate
                                                     select new { d.SanitaryDavObsID, c.Name_Inspector }).FirstOrDefault();

                            if (ID_Name_Inspector != null)
                            {
                                SanitaryDavOb sanitaryDavObToChange = (from c in dbT2.SanitaryDavObs
                                                                       where c.SanitaryDavObsID == ID_Name_Inspector.SanitaryDavObsID
                                                                       select c).FirstOrDefault();

                                if (sanitaryDavObToChange != null)
                                {
                                    sanitaryDavObToChange.Name_Inspector = ID_Name_Inspector.Name_Inspector;
                                }
                            }

                        }

                    }


                    if (sanitaryDonSite.Locator.Substring(0, 2) == "PE")
                    {
                        List<SanitaryDonOb> sanitaryDonObList = (from c in dbT.SanitaryDonObs
                                                                 where c.SanitaryDonSiteID == sanitaryDonSite.SanitaryDonSiteID
                                                                 select c).ToList();

                        // doing Dav
                        SanitaryDavSite sanitaryDavSite = (from c in sanitaryDavSiteList
                                                           where c.Locator == sanitaryDonSite.Locator
                                                           && c.Site == sanitaryDonSite.Site
                                                           select c).FirstOrDefault();

                        if (sanitaryDavSite != null)
                        {
                            List<SanitaryDavOb> sanitaryDavObList = (from c in dbT.SanitaryDavObs
                                                                     where c.SanitaryDavSiteID == sanitaryDavSite.SanitaryDavSiteID
                                                                     select c).ToList();

                            var ID_Name_Inspector = (from c in sanitaryDonObList
                                                     from d in sanitaryDavObList
                                                     where c.ObsDate == d.ObsDate
                                                     select new { d.SanitaryDavObsID, c.Name_Inspector }).FirstOrDefault();

                            if (ID_Name_Inspector != null)
                            {
                                SanitaryDavOb sanitaryDavObToChange = (from c in dbT2.SanitaryDavObs
                                                                       where c.SanitaryDavObsID == ID_Name_Inspector.SanitaryDavObsID
                                                                       select c).FirstOrDefault();

                                if (sanitaryDavObToChange != null)
                                {
                                    sanitaryDavObToChange.Name_Inspector = ID_Name_Inspector.Name_Inspector;
                                }
                            }

                        }

                    }

                    dbT2.SaveChanges();
                }
            }
        }
    }
}
