using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;

using System.Data.OleDb;
using System.IO;
using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet;
using RBI.BUS.BUSExcel;
using RBI.Object.ObjectMSSQL;
using RBI.BUS.BUSMSSQL;
using RBI.PRE.subForm.OutputDataForm;
using DevExpress.XtraSplashScreen;

namespace RBI.PRE.subForm.InputDataForm
{
    public partial class frmImportExcelTank : Form
    {
        public bool ButtonOKClicked { set; get; }
        public frmImportExcelTank()
        {
            InitializeComponent();
        }
        
        #region Parameter
        DevExpress.XtraSpreadsheet.SpreadsheetControl spreadExcel = new SpreadsheetControl();
        string fileName = null;
        string extension = null;
        #endregion
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "Excel 2003 (*.xls)|*.xls|Excel Document (*.xlsx)|*.xlsx|All File(*)|*";
            if (op.ShowDialog() == DialogResult.OK)
            {
                txtPathFileExcel.Text = op.FileName;
            }
        }
        private bool CheckFormatFile()
        {
            IWorkbook workbook = spreadExcel.Document;
            DevExpress.Spreadsheet.Worksheet worksheet = workbook.Worksheets[0];
            bool isCorrect = true;
            if (workbook.Worksheets.Count != 7)
            {
                MessageBox.Show("Format is not correct! Please check again", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return
                    false;
            }
            for (int i = 0; i < 6; i++)
            {
                string sheetName = workbook.Worksheets[i].Name;
                switch (i)
                {
                    case 0:
                        if (sheetName != "Equipment")
                        {
                            MessageBox.Show("Sheet Name " + sheetName + " is not correct!", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            isCorrect = false;
                        }
                        break;
                    case 1:
                        if (sheetName != "Component")
                        {
                            MessageBox.Show("Sheet Name " + sheetName + " is not correct!", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            isCorrect = false;
                        }
                        break;
                    case 2:
                        if (sheetName != "Operating Condition")
                        {
                            MessageBox.Show("Sheet Name " + sheetName + " is not correct!", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            isCorrect = false;
                        }
                        break;
                    case 3:
                        if (sheetName != "Stream")
                        {
                            MessageBox.Show("Sheet Name " + sheetName + " is not correct!", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            isCorrect = false;
                        }
                        break;
                    case 4:
                        if (sheetName != "Material")
                        {
                            MessageBox.Show("Sheet Name " + sheetName + " is not correct!", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            isCorrect = false;
                        }
                        break;
                    default:
                        if (sheetName != "CoatingCladdingLiningInsulation")
                        {
                            MessageBox.Show("Sheet Name " + sheetName + " is not correct!", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            isCorrect = false;
                        }
                        break;
                }
            }
            if (worksheet.Columns.LastUsedIndex <= 31)
            {
                MessageBox.Show("This is Plant Process excel file! Select again", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return isCorrect;
        }
        private void btnFilter_Click(object sender, EventArgs e)
        {
            if (txtPathFileExcel.Text == "")
            {
                MessageBox.Show("Please select a file", "Cortek RBI");
                return;
            }
            fileName = Path.GetFileName(txtPathFileExcel.Text);
            extension = Path.GetExtension(fileName);
            if (extension == ".xls")
            {
                spreadExcel.LoadDocument(txtPathFileExcel.Text, DocumentFormat.Xls);
                if (!CheckFormatFile()) return;
            }
            else if (extension == ".xlsx")
            {
                spreadExcel.LoadDocument(txtPathFileExcel.Text, DocumentFormat.Xlsx);
                if (!CheckFormatFile()) return;
            }
            else
            {
                MessageBox.Show("This file is not supported!", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            btnImport.Enabled = true;
            pictureBox1.Hide();
            label1.Hide();
            spreadExcel.Dock = DockStyle.Fill;
            panelControl1.Controls.Add(spreadExcel);
            btnSave.Enabled = true;
            btnSaveAs.Enabled = true;
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = spreadExcel.Document;
            using (FileStream stream = new FileStream(txtPathFileExcel.Text, FileMode.Create, FileAccess.ReadWrite))
            {
                if (extension == ".xls")
                    workbook.SaveDocument(stream, DocumentFormat.Xls);
                else
                    workbook.SaveDocument(stream, DocumentFormat.Xlsx);
            }
            MessageBox.Show("This file has been saved!", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm2));
            SITES_BUS busSite = new SITES_BUS();
            FACILITY_BUS busFacility = new FACILITY_BUS();
            MANUFACTURER_BUS busManufacture = new MANUFACTURER_BUS();
            DESIGN_CODE_BUS busDesignCode = new DESIGN_CODE_BUS();
            FACILITY_RISK_TARGET_BUS busRiskTarget = new FACILITY_RISK_TARGET_BUS();
            EQUIPMENT_MASTER_BUS busEquipMaster = new EQUIPMENT_MASTER_BUS();
            COMPONENT_MASTER_BUS busCompMaster = new COMPONENT_MASTER_BUS();
            RW_ASSESSMENT_BUS busAss = new RW_ASSESSMENT_BUS();
            RW_EQUIPMENT_BUS busEquip = new RW_EQUIPMENT_BUS();
            RW_COMPONENT_BUS busCom = new RW_COMPONENT_BUS();
            RW_EXTCOR_TEMPERATURE_BUS busExtcor = new RW_EXTCOR_TEMPERATURE_BUS();
            RW_STREAM_BUS busStream = new RW_STREAM_BUS();
            RW_MATERIAL_BUS busMaterial = new RW_MATERIAL_BUS();
            RW_COATING_BUS busCoating = new RW_COATING_BUS();
            RW_INPUT_CA_TANK_BUS busInputCATank = new RW_INPUT_CA_TANK_BUS();
            Bus_PLANT_PROCESS_Excel busExcel = new Bus_PLANT_PROCESS_Excel();
            busExcel.path = txtPathFileExcel.Text;
            //Sites
            List<SITES> lstSite = busExcel.getAllSite();
            foreach(SITES s in lstSite)
            {
                if(!busSite.checkExist(s.SiteName))
                {
                    busSite.add(s);
                }
                else
                {
                    busSite.edit(s);
                }
            }

            //Facility
            List<FACILITY> lstFacility = busExcel.getFacility();
            foreach(FACILITY f in lstFacility)
            {
                if(busFacility.checkExist(f.FacilityName))
                {
                    busFacility.edit(f);
                }
                else
                {
                    busFacility.add(f);
                    int FaciID = busFacility.getLastFacilityID();
                    FACILITY_RISK_TARGET riskTarget = new FACILITY_RISK_TARGET();
                    riskTarget.FacilityID = FaciID;
                    busRiskTarget.add(riskTarget);
                }
            }

            //Manufacture
            List<string> manufacture = busExcel.getAllManufacture();
            for (int i = 0; i < manufacture.Count; i++ )
            {
                MANUFACTURER manu = new MANUFACTURER();
                manu.ManufacturerID = busManufacture.getIDbyName(manufacture[i]);
                manu.ManufacturerName = manufacture[i];
                if(busManufacture.getIDbyName(manufacture[i]) == 0)
                {
                    busManufacture.add(manu);
                }
                else
                {
                    busManufacture.edit(manu);
                }
            }

            //Design Code
            List<string> designCode = busExcel.getAllDesigncode();
            for (int i = 0; i < designCode.Count; i++ )
            {
                DESIGN_CODE design = new DESIGN_CODE();
                design.DesignCodeID = busDesignCode.getIDbyName(designCode[i]);
                design.DesignCode = designCode[i];
                design.DesignCodeApp = "";
                if(design.DesignCodeID == 0)
                {
                    busDesignCode.add(design);
                }
                else
                {
                    busDesignCode.edit(design);
                }
            }

            //Equipment Master
            List<EQUIPMENT_MASTER> lstEquipMaster = busExcel.getEquipmentMaster();
            foreach(EQUIPMENT_MASTER eq in lstEquipMaster)
            {
                if(busEquipMaster.check(eq.EquipmentNumber))
                {
                    busEquipMaster.edit(eq);
                }
                else
                {
                    busEquipMaster.add(eq);
                }
            }

            //Component Master
            List<COMPONENT_MASTER> lstCompMaster = busExcel.getComponentMaster();
            foreach(COMPONENT_MASTER com in lstCompMaster)
            {
                if(busCompMaster.checkExist(com.ComponentNumber))
                {
                    com.ComponentID = busCompMaster.getIDbyName(com.ComponentNumber);
                    busCompMaster.edit(com);
                }
                else
                {
                    busCompMaster.add(com);
                }
            }

            //Rw Assessment
            List<RW_ASSESSMENT> lstAssessment = busExcel.getAssessment();
            List<int> editExcel = new List<int>();
            List<int> addExcel = new List<int>();
            foreach(RW_ASSESSMENT ass in lstAssessment)
            {
                List<int[]> ID_checkAddbyExcel = busAss.getCheckAddExcel_ID(ass.ComponentID, ass.EquipmentID);
                if(ID_checkAddbyExcel.Count != 0)
                {
                    for (int i = 0; i < ID_checkAddbyExcel.Count; i++)
                    {
                        if (ID_checkAddbyExcel[i][0] != 0) //kiem tra xem co phai Assessment nay duoc them tu file Excel ko
                        {
                            ass.ID = ID_checkAddbyExcel[i][1];
                            editExcel.Add(ass.ID);
                            busAss.edit(ass);
                        }
                    }
                }
                else
                {
                    ass.AddByExcel = 1;
                    busAss.add(ass);
                    int assID = busAss.getLastID();
                    addExcel.Add(assID);
                    RW_INPUT_CA_TANK inputCATank = new RW_INPUT_CA_TANK();
                    inputCATank.ID = assID;
                    busInputCATank.add(inputCATank);
                }
            }

            //RW Equipment
            List<RW_EQUIPMENT> lstEquipment = busExcel.getRwEquipmentTank();
            for (int i = 0; i < lstEquipment.Count; i++ )
            {
                if(editExcel.Count != 0)
                {
                    for(int j = 0; j < editExcel.Count; j++)
                    {
                        if(lstEquipment[i].ID == editExcel[j])
                        {
                            busEquip.edit(lstEquipment[i]);
                        }
                    }
                }
                if (addExcel.Count != 0)
                {
                    
                    for (int j = 0; j < addExcel.Count; j++)
                    {
                        if (lstEquipment[i].ID == addExcel[j])
                        {
                            busEquip.add(lstEquipment[i]);
                        }
                    }
                }
            }

            //RW Component
            List<RW_COMPONENT> lstComponent = busExcel.getRwComponentTank();
            for (int i = 0; i < lstComponent.Count; i++)
            {
                if (editExcel.Count != 0)
                {
                    for (int j = 0; j < editExcel.Count; j++)
                    {
                        if (lstComponent[i].ID == editExcel[j])
                        {
                            busCom.edit(lstComponent[i]);
                        }
                    }
                }
                if (addExcel.Count != 0)
                {
                    for (int j = 0; j < addExcel.Count; j++)
                    {
                        if (lstComponent[i].ID == addExcel[j])
                        {
                            busCom.add(lstComponent[i]);
                        }
                    }
                }
            }

            //RW Extcor temperature
            List<RW_EXTCOR_TEMPERATURE> lstExtcor = busExcel.getRwExtTemp();
            for (int i = 0; i < lstExtcor.Count; i++)
            {
                if (editExcel.Count != 0)
                {
                    for (int j = 0; j < editExcel.Count; j++)
                    {
                        if (lstExtcor[i].ID == editExcel[j])
                        {
                            busExtcor.edit(lstExtcor[i]);
                        }
                    }
                }
                if (addExcel.Count != 0)
                {
                    for (int j = 0; j < addExcel.Count; j++)
                    {
                        if (lstExtcor[i].ID == addExcel[j])
                        {
                            busExtcor.add(lstExtcor[i]);
                        }
                    }
                }
            }

            //RW Stream
            List<RW_STREAM> lstStream = busExcel.getRwStreamTank();
            for (int i = 0; i < lstStream.Count; i++)
            {
                if (editExcel.Count != 0)
                {
                    for (int j = 0; j < editExcel.Count; j++)
                    {
                        if (lstStream[i].ID == editExcel[j])
                        {
                            busStream.edit(lstStream[i]);
                        }
                    }
                }
                if (addExcel.Count != 0)
                {
                    for (int j = 0; j < addExcel.Count; j++)
                    {
                        if (lstStream[i].ID == addExcel[j])
                        {
                            busStream.add(lstStream[i]);
                        }
                    }
                }
            }

            //RW Material
            List<RW_MATERIAL> lstMaterial = busExcel.getRwMaterialTank();
            for (int i = 0; i < lstMaterial.Count; i++)
            {
                if (editExcel.Count != 0)
                {
                    for (int j = 0; j < editExcel.Count; j++)
                    {
                        if (lstMaterial[i].ID == editExcel[j])
                        {
                            busMaterial.edit(lstMaterial[i]);
                        }
                    }
                }
                if (addExcel.Count != 0)
                {
                    for (int j = 0; j < addExcel.Count; j++)
                    {
                        if (lstMaterial[i].ID == addExcel[j])
                        {
                            busMaterial.add(lstMaterial[i]);
                        }
                    }
                }
            }

            //RW Coating
            List<RW_COATING> lstCoating = busExcel.getRwCoating();
            for (int i = 0; i < lstCoating.Count; i++)
            {
                if (editExcel.Count != 0)
                {
                    for (int j = 0; j < editExcel.Count; j++)
                    {
                        if (lstCoating[i].ID == editExcel[j])
                        {
                            busCoating.edit(lstCoating[i]);
                        }
                    }
                }
                if (addExcel.Count != 0)
                {
                    for (int j = 0; j < addExcel.Count; j++)
                    {
                        if (lstCoating[i].ID == addExcel[j])
                        {
                            busCoating.add(lstCoating[i]);
                        }
                    }
                }
            }
            ButtonOKClicked = true;
            SplashScreenManager.CloseForm();
            MessageBox.Show("All data have been saved! You need to add Risk Target in Facility!", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            this.Close();
        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        //<read excel file with spreadsheet control>

        //private void readDataFromSheet()
        //{
        //    IWorkbook workbook = spreadExcel.Document;
        //    try
        //    {
        //        //sheet equipment
        //        DevExpress.Spreadsheet.Worksheet worksheet0 = workbook.Worksheets[0];
        //        sheetEquipment = new string[33];
        //        sheetEquipment[0] = "Sheet Equipment";
        //        for (int i = 1; i < 33; i++)
        //        {
        //            sheetEquipment[i] = worksheet0.Cells[1, i-1].Value.ToString();
        //            Debug.WriteLine("Equipment " + sheetEquipment[i]);
        //        }
        //        //sheet component
        //        DevExpress.Spreadsheet.Worksheet worksheet1 = workbook.Worksheets[1];
        //        sheetComponent = new string[33];
        //        sheetComponent[0] = "Sheet Component";
        //        for (int i = 1; i < 33; i++)
        //        {
        //            sheetComponent[i] = worksheet1.Cells[1, i-1].Value.ToString();
        //            Debug.WriteLine("Component " + sheetComponent[i]);
        //        }
        //        //sheet operating condition
        //        DevExpress.Spreadsheet.Worksheet worksheet2 = workbook.Worksheets[2];
        //        sheetOperatingCondition = new string[19];
        //        sheetOperatingCondition[0] = "Sheet Operating Condition";
        //        for (int i = 1; i < 19; i++)
        //        {
        //            sheetOperatingCondition[i] = worksheet2.Cells[1, i-1].Value.ToString();
        //            Debug.WriteLine("Operating Condition " + sheetOperatingCondition[i]);
        //        }
        //        //sheet Stream
        //        DevExpress.Spreadsheet.Worksheet worksheet3 = workbook.Worksheets[3];
        //        sheetStream = new string[23];
        //        sheetStream[0] = "Sheet Stream";
        //        for (int i = 1; i < 23; i++)
        //        {
        //            sheetStream[i] = worksheet3.Cells[1, i-1].Value.ToString();
        //            Debug.WriteLine("Stream " + sheetStream[i]);
        //        }
        //        //sheet Material
        //        DevExpress.Spreadsheet.Worksheet worksheet4 = workbook.Worksheets[4];
        //        sheetMaterial = new string[23];
        //        sheetMaterial[0] = "Sheet Material";
        //        for (int i = 1; i < 23; i++)
        //        {
        //            sheetMaterial[i] = worksheet4.Cells[1, i-1].Value.ToString();
        //            Debug.WriteLine("Material " + sheetMaterial[i]);
        //        }
        //        //sheet Coating
        //        DevExpress.Spreadsheet.Worksheet worksheet5 = workbook.Worksheets[5];
        //        sheetCoating = new string[16];
        //        sheetCoating[0] = "Sheet Coating";
        //        for (int i = 1; i < 16; i++)
        //        {
        //            sheetCoating[i] = worksheet5.Cells[1, i-1].Value.ToString();
        //            Debug.WriteLine("Coating " + sheetCoating[i]);
        //        }

        //    }
        //    catch(Exception e)
        //    {
        //        MessageBox.Show("Error Read Excel File!" + e.ToString(), "Cortek");
        //    }
        //}
        //public RW_EQUIPMENT getDataEquipment()
        //{
        //    RW_EQUIPMENT eq = new RW_EQUIPMENT();
        //    eq.AdminUpsetManagement = sheetEquipment[12] == "True" ? 1 : 0;
        //    eq.ContainsDeadlegs = sheetEquipment[13] == "True" ? 1 : 0;
        //    eq.CyclicOperation = sheetEquipment[15] == "True" ? 1 : 0;
        //    eq.HighlyDeadlegInsp = sheetEquipment[14] == "True" ? 1 : 0;
        //    eq.DowntimeProtectionUsed = sheetEquipment[16] == "True" ? 1 : 0;
        //    eq.ExternalEnvironment = sheetEquipment[28];
        //    eq.HeatTraced = sheetEquipment[18] == "True" ? 1 : 0;
        //    eq.InterfaceSoilWater = sheetEquipment[20] == "True" ? 1 : 0;
        //    eq.LinerOnlineMonitoring = sheetEquipment[25] == "True" ? 1 : 0;
        //    eq.MaterialExposedToClExt = sheetEquipment[24] == "True" ? 1 : 0;
        //    eq.MinReqTemperaturePressurisation = sheetEquipment[22] != "" ? float.Parse(sheetEquipment[22]) : 0;
        //    eq.OnlineMonitoring = sheetEquipment[30];
        //    eq.PresenceSulphidesO2 = sheetEquipment[26] == "True" ? 1 : 0;
        //    eq.PresenceSulphidesO2Shutdown = sheetEquipment[27] == "True" ? 1 : 0;
        //    eq.PressurisationControlled = sheetEquipment[21] == "True" ? 1 : 0;
        //    eq.PWHT = sheetEquipment[19] == "True" ? 1 : 0;
        //    eq.SteamOutWaterFlush = sheetEquipment[17] == "True" ? 1 : 0;
        //    eq.ManagementFactor = sheetEquipment[32] != "" ? float.Parse(sheetEquipment[32]) : 0;
        //    eq.ThermalHistory = sheetEquipment[29];
        //    eq.YearLowestExpTemp = sheetEquipment[23] == "True" ? 1 : 0;
        //    eq.Volume = sheetEquipment[31] != "" ? float.Parse(sheetEquipment[31]) : 0;
        //    eq.CommissionDate = Convert.ToDateTime(sheetEquipment[8]);
        //    Debug.WriteLine("commission date " + eq.CommissionDate);
        //    return eq;
        //}
        //public RW_ASSESSMENT getAssessmentDate()
        //{
        //    RW_ASSESSMENT ass = new RW_ASSESSMENT();
        //    ass.AssessmentDate = Convert.ToDateTime(sheetComponent[8]);
        //    Debug.WriteLine("assessment date " + ass.AssessmentDate);
        //    return ass;
        //}
        //public RW_COMPONENT getDataComponent()
        //{
        //    RW_COMPONENT comp = new RW_COMPONENT();
        //    comp.NominalDiameter = sheetComponent[10] == "" ? float.Parse(sheetComponent[10]) : 0;
        //    comp.NominalThickness = sheetComponent[11] == "" ? float.Parse(sheetComponent[11]) : 0;
        //    comp.CurrentThickness = sheetComponent[12] == "" ? float.Parse(sheetComponent[12]) : 0;
        //    comp.MinReqThickness = sheetComponent[13] == "" ? float.Parse(sheetComponent[13]) : 0;
        //    comp.CurrentCorrosionRate = sheetComponent[14] == "" ? float.Parse(sheetComponent[14]) : 0;
        //    comp.BranchDiameter = sheetComponent[24];
        //    comp.BranchJointType = sheetComponent[25];
        //    comp.BrinnelHardness = sheetComponent[21];
        //    comp.CracksPresent = sheetComponent[17] == "True" ? 1 : 0;
        //    comp.ComplexityProtrusion = sheetComponent[22];
        //    return comp;
        //}
        //public RW_COATING getDataCoating()
        //{
        //    RW_COATING coat = new RW_COATING();
        //    try
        //    {
        //        coat.ExternalCoating = sheetCoating[3] == "True" ? 1 : 0;
        //        coat.ExternalInsulation = sheetCoating[12] == "True" ? 1 : 0;
        //        coat.InternalCladding = sheetCoating[7] == "True" ? 1 : 0;
        //        coat.InternalCoating = sheetCoating[2] == "True" ? 1 : 0;
        //        coat.ExternalCoatingQuality = sheetCoating[5];
        //        coat.ExternalInsulationType = sheetCoating[14];
        //        coat.InsulationContainsChloride = sheetCoating[13] == "True" ? 1 : 0;
        //        coat.InternalLinerCondition = sheetCoating[11];
        //        coat.InternalLinerType = sheetCoating[10];
        //        coat.InternalLining = sheetCoating[9] == "True" ? 1 : 0;
        //        coat.CladdingCorrosionRate = sheetCoating[8] != "" ? float.Parse(sheetCoating[8]) : 0;
        //        coat.SupportConfigNotAllowCoatingMaint = sheetCoating[6] == "True" ? 1 : 0;
        //    }
        //    catch(Exception e)
        //    {
        //        MessageBox.Show(e.ToString());
        //    }
        //    return coat;
        //}
        //public RW_STREAM getDataStream()
        //{
        //    RW_STREAM stream = new RW_STREAM();
        //    stream.AmineSolution = sheetStream[13];
        //    stream.AqueousOperation = sheetStream[14] == "True" ? 1 : 0;
        //    stream.AqueousShutdown = sheetStream[15] == "True" ? 1 : 0;
        //    stream.ToxicConstituent = sheetStream[11] == "True" ? 1 : 0;
        //    stream.Caustic = sheetStream[20] == "True" ? 1 : 0;
        //    stream.Chloride = sheetStream[6] != "" ? float.Parse(sheetStream[6]) : 0;
        //    stream.CO3Concentration = sheetStream[7] != "" ? float.Parse(sheetStream[7]) : 0;
        //    stream.Cyanide = sheetStream[18] == "True" ? 1 : 0;
        //    stream.ExposedToGasAmine = sheetStream[10] == "True" ? 1 : 0;
        //    stream.ExposedToSulphur = sheetStream[21] == "True" ? 1 : 0;
        //    stream.ExposureToAmine = sheetStream[12];
        //    stream.H2S = sheetStream[16] == "True" ? 1 : 0;
        //    stream.H2SInWater = sheetStream[8] != "" ? float.Parse(sheetStream[8]) : 0;
        //    stream.Hydrogen = sheetStream[19] == "True" ? 1 : 0;
        //    stream.MaterialExposedToClInt = sheetStream[22] == "True" ? 1 : 0;
        //    stream.NaOHConcentration = sheetStream[4] != "" ? float.Parse(sheetStream[4]) : 0;
        //    stream.ReleaseFluidPercentToxic = sheetStream[5] != "" ? float.Parse(sheetStream[5]) : 0;
        //    stream.WaterpH = sheetStream[9] != "" ? float.Parse(sheetStream[9]) : 0;
        //    stream.Hydrofluoric = sheetStream[17] == "True" ? 1 : 0;
        //    return stream;

        //}
        //public RW_MATERIAL getDataMaterial()
        //{
        //    RW_MATERIAL ma = new RW_MATERIAL();
        //    ma.MaterialName = sheetMaterial[1];
        //    ma.DesignPressure = sheetMaterial[2] != "" ? float.Parse(sheetMaterial[2]) : 0;
        //    ma.DesignTemperature = sheetMaterial[3] != "" ? float.Parse(sheetMaterial[3]) : 0;
        //    ma.MinDesignTemperature = sheetMaterial[4] != "" ? float.Parse(sheetMaterial[4]) : 0;
        //    ma.BrittleFractureThickness = sheetMaterial[6] != "" ? float.Parse(sheetMaterial[6]) : 0;
        //    ma.CorrosionAllowance = sheetMaterial[8] != "" ? float.Parse(sheetMaterial[8]) : 0;
        //    //if(tankBottom) -> hide txtSigmaPhase
        //    ma.SigmaPhase = sheetMaterial[9] != "" ? float.Parse(sheetMaterial[8]) : 0;
        //    ma.SulfurContent = sheetMaterial[15];
        //    ma.HeatTreatment = sheetMaterial[16];
        //    ma.ReferenceTemperature = sheetMaterial[5] != "" ? float.Parse(sheetMaterial[5]) : 0;
        //    ma.PTAMaterialCode = sheetMaterial[20];
        //    ma.HTHAMaterialCode = sheetMaterial[18];
        //    ma.IsPTA = sheetMaterial[19] == "True" ? 1 : 0;
        //    ma.IsHTHA = sheetMaterial[17] == "True" ? 1 : 0;
        //    ma.Austenitic = sheetMaterial[11] == "True" ? 1 : 0;
        //    ma.Temper = sheetMaterial[12] == "True" ? 1 : 0;
        //    ma.CarbonLowAlloy = sheetMaterial[10] == "True" ? 1 : 0;
        //    ma.NickelBased = sheetMaterial[13] == "True" ? 1 : 0;
        //    ma.ChromeMoreEqual12 = sheetMaterial[14] == "True" ? 1 : 0;
        //    ma.AllowableStress = sheetMaterial[7] != "" ? float.Parse(sheetMaterial[7]) : 0;
        //    //ma.CostFactor = mate[20] != "" ? float.Parse(mate[20]) : 0;
        //    return ma;
        //}
        //public RW_STREAM getDataOperating()
        //{
        //    RW_STREAM str = new RW_STREAM();
        //    str.FlowRate = sheetOperatingCondition[8] != "" ? float.Parse(sheetOperatingCondition[8]) : 0;
        //    str.MaxOperatingPressure = sheetOperatingCondition[5] != "" ? float.Parse(sheetOperatingCondition[5]) : 0;
        //    str.MinOperatingPressure = sheetOperatingCondition[6] != "" ? float.Parse(sheetOperatingCondition[6]) : 0;
        //    str.MaxOperatingTemperature = sheetOperatingCondition[2] != "" ? float.Parse(sheetOperatingCondition[2]) : 0;
        //    str.MinOperatingTemperature = sheetOperatingCondition[3] != "" ? float.Parse(sheetOperatingCondition[3]) : 0;
        //    str.CriticalExposureTemperature = sheetOperatingCondition[4] != "" ? float.Parse(sheetOperatingCondition[4]) : 0;
        //    str.H2SPartialPressure = sheetOperatingCondition[7] != "" ? float.Parse(sheetOperatingCondition[7]) : 0;
        //    str.CUI_PERCENT_1 = sheetOperatingCondition[9] != "" ? float.Parse(sheetOperatingCondition[9]) : 0;
        //    str.CUI_PERCENT_2 = sheetOperatingCondition[10] != "" ? float.Parse(sheetOperatingCondition[10]) : 0;
        //    str.CUI_PERCENT_3 = sheetOperatingCondition[11] != "" ? float.Parse(sheetOperatingCondition[11]) : 0;
        //    str.CUI_PERCENT_4 = sheetOperatingCondition[12] != "" ? float.Parse(sheetOperatingCondition[12]) : 0;
        //    str.CUI_PERCENT_5 = sheetOperatingCondition[13] != "" ? float.Parse(sheetOperatingCondition[13]) : 0;
        //    str.CUI_PERCENT_6 = sheetOperatingCondition[14] != "" ? float.Parse(sheetOperatingCondition[14]) : 0;
        //    str.CUI_PERCENT_7 = sheetOperatingCondition[15] != "" ? float.Parse(sheetOperatingCondition[15]) : 0;
        //    str.CUI_PERCENT_8 = sheetOperatingCondition[16] != "" ? float.Parse(sheetOperatingCondition[16]) : 0;
        //    str.CUI_PERCENT_9 = sheetOperatingCondition[17] != "" ? float.Parse(sheetOperatingCondition[17]) : 0;
        //    str.CUI_PERCENT_10 = sheetOperatingCondition[18] != "" ? float.Parse(sheetOperatingCondition[18]) : 0;
        //    return str;
        //}

        //</read excel file with spreadsheet control>
        private void btnSaveAs_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = spreadExcel.Document;
            SaveFileDialog op = new SaveFileDialog();
            op.Filter = "Excel 2003 (*.xls)|*.xls|Excel Document (*.xlsx)|*.xlsx|All File(*)|*";
            op.Title = "Save Inspection History File";
            op.ShowDialog();
            String pathFile = op.FileName;
            String exten = Path.GetExtension(pathFile);
            if (pathFile != "")
            {
                try
                {
                    using (FileStream stream = new FileStream(pathFile, FileMode.Create, FileAccess.ReadWrite))
                    {
                        if (exten == ".xls")
                            workbook.SaveDocument(stream, DocumentFormat.Xls);
                        else
                            workbook.SaveDocument(stream, DocumentFormat.Xlsx);
                    }
                    MessageBox.Show("This file has been saved!", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                catch
                {
                    MessageBox.Show("Save file error!", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

    }
}

