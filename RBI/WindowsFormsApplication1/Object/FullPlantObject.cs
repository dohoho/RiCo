﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RBI.Object
{
    class FullPlantObject
    {
        public String PLANT { set; get; }
        public String Unit { set; get; }
        public String EquipNum { set; get; }
        public String EquipDescrip { set; get; }
        public String EquipType { set; get; }
        public String SubComponent { set; get; }
        public String SubComponentDescrip { set; get; }
        public String MOC { set; get; }
        public String LMOC { set; get; }
        public double HeightLength { set; get; }
        public double Diameter { set; get; }
        public double NominalThick { set; get; }
        public double CA { set; get; }
        public double DesignPressure { set; get; }
        public double DesignTemp { set; get; }
        public double OPPressInlet { set; get; }
        public double OPTempInlet { set; get; }
        public double OPPressOutlet { set; get; }
        public double OPTempOutlet { set; get; }
        public double TestPress { set; get; }
        public double MDMT { set; get; }
        public Boolean InService { set; get; }
        public String ServiceDate { set; get; }
        public String LastIsnpDate { set; get; }
        public Boolean LDTBXH { set; get; }
        public Boolean Insulated { set; get; }
        public Boolean PWHT { set; get; }
        public String InsulationType { set; get; }
        public String OperatingState { set; get; }
        public double InventoryLiquid { set; get; }
        public double InventoryVapor { set; get; }
        public double InventoryTotal { set; get; }
        public Boolean ConfidentInstream { set; get; }
        public Boolean VaporDensityLessAir { set; get; }
        public Boolean CorrosionInhibitor { set; get; }
        public Boolean PrequentFeedChange { set; get; }
        public double MajorChemicals { set; get; }
        public String Contaminant { set; get; }
        public Boolean OnlineMonitor { set; get; }
        public Boolean CathodicProtection { set; get; }
        public Boolean CorrosionMonitor { set; get; }
        public Boolean OHCalibUpDate { set; get; }
        public double DistFromFacility { set; get; }
        public double EquipCount { set; get; }
        public double HAZOPRate { set; get; }
        public double PersonDensity { set; get; }
        public double MitigationEquip { set; get; }
        public double EnvRate { set; get; }
        public String InsTechUsed { set; get; }
        public String EquidModification_Repair { set; get; }
        public String InspFinding { set; get; }
        public double VaporDensity { set; get; }
        public double LiquidDensity { set; get; }
        public double Vapor { set; get; }
        public double Liquid { set; get; }
        public String HMBPFDNum { set; get; }
        public String PIDNum { set; get; }
        public String Service { set; get; }
        public String HMBStream { set; get; }
        public Boolean CrackPresent { set; get; }
        public Boolean ProtectedBarrier { set; get; }
        public String ComponentType { set; get; }
        public String LastCrackingInspDate { set; get; }
        public Boolean InternalLiner { set; get; }
        public String CatalogThin { set; get; }
        public int NoInsp { set; get; }
        public String CheckThin { set; get; }
        public Boolean Cladding { set; get; }
        public int Fom { set; get; }
        public int Fip { set; get; }
        public int Fdl { set; get; }
        public int Fwd { set; get; }
        public int Fam { set; get; }
        public double Fsm { set; get; }
        public double CorrosionRateMetal { set; get; }
        public double CorrosionRateCladding { set; get; }
        public double MinimumThick { set; get; }
        public double ThickBaseMetal { set; get; }
        public String LinningType { set; get; }
        public int Flc { set; get; }
        public int YearInservice { set; get; }
        public String LevelCaustic { set; get; }
        public String CatalogCaustic { set; get; }
        public String LevelAmine { set; get; }
        public String CatalogAmine { set; get; }
        public String CatalogSulfide { set; get; }
        public double pH { set; get; }
        public double Sulfide_ppm { set; get; }
        // No PWHT Do Cung Brinell
        public double NoPWHT { set; get; }
        public String HIC_H2S_Catalog { set; get; }
        public double H2S_ppm { set; get; }
        public String CacbonateCatalog { set; get; }
        public double Cacbonate_ppm { set; get; }
        public String CatalogPTA { set; get; }
        public String Materials { set; get; }
        public String HeatTreatment { set; get; }
        public String CatalogCLSCC { set; get; }
        public double TempPH { set; get; }
        public double Clo_ppm { set; get; }
        public String Catalog_HF { set; get; }
        public Boolean HFpresent { set; get; }
        public double BrinellHardness { set; get; }
        public String Catalog_HIC_HF { set; get; }
        public double SulfurPercent { set; get; }
        public String Catalog_External { set; get; }
        public String ExternalDriver { set; get; }
        public String CoatQuality { set; get; }
        public String CatalogCUI { set; get; }
        public String DriverCUI { set; get; }
        public double CorrosionRateCUI { set; get; }
        public String Complexity { set; get; }
        public String Insulation { set; get; }
        public Boolean AllowConfig { set; get; }
        public Boolean EnterSoil { set; get; }
        public String InsulationTypeCUI { set; get; }
        public String CatalogExtCLSCC { set; get; }
        public String DriverExtCLSCC { set; get; }
        public String PipingComplexity { set; get; }
        public String InsulationCondition { set; get; }
        public String HTHA_Catalog { set; get; }
        public int AgeHTHA { set; get; }
        public double TempHTHA { set; get; }
        public double PH2 { set; get; }
        public double TempMinBrittle { set; get; }
        public double TempUpsetBrittle { set; get; }
        public double NBP { set; get; }
        public double TempImpact { set; get; }
        public String MaterialCurve { set; get; }
        public Boolean LowTemp { set; get; }
        public double SCE { set; get; }
        public double ReferenceTemp { set; get; }
        public double TempMin885 { set; get; }
        public Boolean BrittleCheck { set; get; }
        public double TempShut { set; get; }
        public double PercentSigma { set; get; }
        public String NoFailure { set; get; }
        public String SeverityVibration { set; get; }
        public double NoWeek { set; get; }
        public String CyclicType { set; get; }
        public String CorrectAction { set; get; }
        public double ToTalPiping { set; get; }
        public String TypeOfPiping { set; get; }
        public String PipeCondition { set; get; }
        public double BranchDiametter { set; get; }
        public String Fluid { set; get; }
        public String MaterialsCA { set; get; }
        public String FluidPhase { set; get; }
        public String FluidType { set; get; }
        public String ReleaseFluid { set; get; }
        public double StoredPressure { set; get; }
        public double AtmosphericPressure { set; get; }
        public double StoredTemp { set; get; }
        public double AtmosphericTemp { set; get; }
        public double Reynol { set; get; }
        public int MitigationSystem { set; get; }
        public String ToxicMaterialsLV1 { set; get; }
        public double ToxicPercent { set; get; }
        public String ReleaseDuration { set; get; }
        public String NonToxic_NonFlammable { set; get; }
        public double OutageMultiplier { set; get; }
        public double ProductionCost { set; get; }
        public double InjuryCost { set; get; }
        public double EnvCost { set; get; }
        public double EquipmentCost { set; get; }
        public String PoolFireType { set; get; }
        public double MassFractionLiquid { set; get; }
        public double FractionFuild { set; get; }
        public double BubblePointTemp { set; get; }
        public double DewPointVapor { set; get; }
        public double TimeSteady { set; get; }
        public double SpecificHeat { set; get; }
        public double MassFrammableVapor { set; get; }
        public double MassFract { set; get; }
        public double VolumeLiquid { set; get; }
        public double BubblePointPress { set; get; }
        public double WindSpeed { set; get; }
        public String AreaType { set; get; }
        public double GroundTemp { set; get; }
        public String AmbientCondition { set; get; }
        public double Humidity { set; get; }
        public double MoleFract { set; get; }
        public String ToxicComponent { set; get; }
        public String Criteria { set; get; }
        public double GradeLevelCloud { set; get; }
        public String RepresentFluid { set; get; }
        public double MoleFlash { set; get; }
        public double MaximumFillHeight { set; get; }
        public String ReleaseHoleSize { set; get; }
        public int ShellCourse { set; get; }
        public double CHT { set; get; }
        public String EnvironSensitivity { set; get; }
        public double P_dike { set; get; }
        public double P_onsite { set; get; }
        public double P_offsite { set; get; }
        public String Tank_type { set; get; }
        public double SoilHydraulic { set; get; }
        public double Distance { set; get; }
        public double Fc { set; get; }
        public double OverPress { set; get; }
        public double MAWP { set; get; }
        public int Fenv { set; get; }
        public Boolean CheckPass { set; get; }
        public String CatalogRelief { set; get; }
        public String FluidSeverityPoF { set; get; }
        public String WelbullPoF { set; get; }
        public String FluidSeverityLeak { set; get; }
        public String WelbullLeak { set; get; }
        public double TotalDemand { set; get; }
        public double Fs { set; get; }
        public Boolean IsLeak { set; get; }
        public String LevelLeak { set; get; }
        public double RateCapacity { set; get; }
        public double TimeIsolate { set; get; }
        public double FluidCost { set; get; }
        public double PRDinletSize { set; get; }
        public String PRDType { set; get; }
        public double Fr { set; get; }
        public double NoDay { set; get; }
        public Boolean IgnoreLeak { set; get; }
        public double RateReduct { set; get; }
        public double MainCost { set; get; }
        public double Fms { set; get; }
        public String DetectionType { set; get; }
        public String IsolationType { set; get; }
    }
}
