using RBI.BUS.BUSMSSQL;
using RBI.Object.ObjectMSSQL;
using RBI.Object.ObjectMSSQL_CAL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RBI.DAL.MSSQL_CAL;

namespace RBI.BUS.BUSMSSQL_CAL
{
    class MSSQL_CA_CAL
    {
        API_COMPONENT_TYPE_BUS API_COMPONENT_BUS = new API_COMPONENT_TYPE_BUS();
        MSSQL_RBI_CAL_ConnUtils DAL_CAL = new MSSQL_RBI_CAL_ConnUtils();
        public float MATERIAL_COST { set; get; }
        public float NominalDiameter { set; get; }
        public String FLUID { set; get; }
        public String FLUID_PHASE { set; get; }
        public String API_COMPONENT_TYPE_NAME { set; get; }
        public String DETECTION_TYPE { set; get; }
        public String ISULATION_TYPE { set; get; }
        public float STORED_PRESSURE { set; get; }
        public float ATMOSPHERIC_PRESSURE { set; get; }//MPa
        public float STORED_TEMP { set; get; }//K
        public float MASS_INVERT { set; get; }
        public float MASS_COMPONENT { set; get; }
        public String MITIGATION_SYSTEM { set; get; }// mitigation factor
        public float TOXIC_PERCENT { set; get; }// mfrac_toxic
        public String RELEASE_DURATION { set; get; }// duration time
        public float PRODUCTION_COST { set; get; }
        public float INJURE_COST { set; get; }
        public float ENVIRON_COST { set; get; }
        public float PERSON_DENSITY { set; get; }
        public float EQUIPMENT_COST { set; get; }
        private String TOXIC_PHASE { set; get; }
        public String FC_Category(float fc)
        {
            if (fc <= 10000)
                return "A";
            else if (fc <= 100000)
                return "B";
            else if (fc <= 1000000)
                return "C";
            else if (fc <= 10000000)
                return "D";
            else
                return "E";
        }
        public String CA_Category(float ca)
        {
            if (ca <= 9.29)
                return "A";
            else if (ca <= 92.9)
                return "B";
            else if (ca <= 279)
                return "C";
            else if (ca <= 929)
                return "D";
            else
                return "E";
        }
        private String TYPE_FLUID()
        {
            String API_TYPE = null;
            if (FLUID == null || FLUID == "")
                API_TYPE = "TYPE 0";
            else
            {
                switch (FLUID)
                {
                    case "C1-C2":
                    case "C13-C16":
                    case "C17-C25":
                    case "C25+":
                    case "C3-C4":
                    case "C5":
                    case "C6-C8":
                    case "C9-C12":
                    case "Acid":
                    case "AlCl3":
                    case "H2":
                    case "H2S":
                    case "HCl":
                    case "HF":
                    case "Nitric Acid":
                    case "NO2":
                    case "Phosgene":
                    case "Pyrophoric":
                    case "Steam":
                    case "TDI":
                    case "Water":
                        API_TYPE = "TYPE 0";
                        break;
                    case "CO":
                    case "DEE":
                    case "EE":
                    case "EEA":
                    case "EG":
                    case "EO":
                    case "Methanol":
                    case "PO":
                    case "Styrene":
                        API_TYPE = "TYPE 1";
                        break;
                    default:
                        API_TYPE = "TYPE 0";
                        break;
                }
            }
            return API_TYPE;
        }
        private API_COMPONENT_TYPE GET_DATA_API_COM()
        {
            return API_COMPONENT_BUS.getData(API_COMPONENT_TYPE_NAME);
        }

        public String GET_RELEASE_PHASE()
        {
            return DAL_CAL.GET_RELEASE_PHASE(FLUID);
        }
        // Step 2 release hole size
        private float d_n(int i)
        {
            float dn = 0;
            switch(API_COMPONENT_TYPE_NAME)
            {
                case "TANKBOTTOM":
                    if (i == 1)
                    {
                        if (PREVENTION_BARRIER)
                            dn = 3.175f;
                        else
                            dn = 12.7f;
                    }
                    else if (i == 2)
                        dn = 0;
                    else if (i == 3)
                        dn = 0;
                    else if (i == 4 && PREVENTION_BARRIER)
                        dn = 250 * TANK_DIAMETER;
                    else
                        dn = 0;
                    break;
                case "COURSE-1":
                case "COURSE-2":
                case "COURSE-3":
                case "COURSE-4":
                case "COURSE-5":
                case "COURSE-6":
                case "COURSE-7":
                case "COURSE-8":
                case "COURSE-9":
                case "COURSE-10":
                    if (i == 1)
                        dn = 3.175f;
                    else if (i == 2)
                        dn = 6.35f;
                    else if (i == 3)
                        dn = 50.8f;
                    else
                        dn = 250 * TANK_DIAMETER;
                    break;
                default:
                    if (i == 1)
                        dn = 6.4f;
                    else if (i == 2)
                        dn = 25;
                    else if (i == 3)
                        dn = 102;
                    else
                        dn = (float)Math.Min(NominalDiameter, 406);
                    break;
            }
            return dn;
        }
        private float a_n(int i)
        {
            return (float)Math.Round(Math.PI * Math.Pow(d_n(i), 2) / 4, 2);
        }
        private float C_P()
        {
            float[] data = DAL_CAL.GET_TBL_52(FLUID);
            float CP_C2 = (float)Math.Round((data[6] / STORED_TEMP) / (Math.Sinh(data[6] / STORED_TEMP)), 2);
            float CP_E2 = (float)Math.Round((data[8] / STORED_TEMP) / (Math.Cosh(data[8] / STORED_TEMP)), 2);
            if (data[3] == 1)
                return data[4] + data[5] * STORED_TEMP + data[6] * (float)Math.Pow(STORED_TEMP, 2) + data[7] * (float)Math.Pow(STORED_TEMP, 3);
            else if (data[3] == 2)
                return data[4] + data[5] * CP_C2 * CP_C2 + data[6] * CP_E2 * CP_E2;
            else if (data[3] == 3)
                return data[4] + data[5] * STORED_TEMP + data[6] * (float)Math.Pow(STORED_TEMP, 2) + data[7] * (float)Math.Pow(STORED_TEMP, 3) + data[8] * (float)Math.Pow(STORED_TEMP, 4);
            else
                return 0;
        }
        private float W_n(int i)
        {
            float[] data = DAL_CAL.GET_TBL_52(FLUID);
            float W_n;
            float an = a_n(i);
            float k;
            float m_w = data[0];
            float p_trans;
            float gc = 1;
            if (FLUID_PHASE == "Liquid" || FLUID_PHASE == "Two-phase" || FLUID_PHASE == "Powder")
            {
                W_n = (float)Math.Round(0.61 * 1 * data[1] * 16.02 * an * Math.Sqrt(2 * gc * Math.Abs(STORED_PRESSURE - ATMOSPHERIC_PRESSURE) / (data[1]) * 16.02) / (DAL_CAL.GET_TBL_3B21(1)), 2);
            }
            else
            {
                float R = 8.314f;
                k = (C_P() / (C_P() - R));
                p_trans = (float)Math.Round(ATMOSPHERIC_PRESSURE * Math.Pow((k + 1) / 2, k / (k - 1)), 2);
                if (STORED_PRESSURE > p_trans)
                {
                    float x = (float)((k * m_w * gc / (R * STORED_TEMP)) * Math.Pow(2 / (k + 1), (k + 1) / (k - 1)));
                    W_n = (float)Math.Round(0.9 * an * STORED_PRESSURE * Math.Sqrt(Math.Abs(x)) / (DAL_CAL.GET_TBL_3B21(2)), 2);
                }
                else
                {
                    float x = (float)Math.Round((m_w * gc / (R * STORED_TEMP)) * ((2 * k) / (k - 1)) * Math.Pow(ATMOSPHERIC_PRESSURE / STORED_PRESSURE, 2 / k) * (1 - Math.Pow(ATMOSPHERIC_PRESSURE / STORED_PRESSURE, (k - 1) / k)), 2);
                    W_n = (float)Math.Round(0.9 * an * STORED_PRESSURE * Math.Sqrt(Math.Abs(x)) / (DAL_CAL.GET_TBL_3B21(2)), 2);
                }
            }
            return W_n;
        }
        private float W_max8()
        {
            float[] data = DAL_CAL.GET_TBL_52(FLUID);
            float W_max8 = 0;
            float an = 32450;
            float k = 0;
            float mw = data[0];
            float p_trans = 0;
            float gc = 1;
            if (FLUID_PHASE == "Liquid" || FLUID_PHASE == "Two-phase" || FLUID_PHASE == "Powder")
            {
                W_max8 = (float)Math.Round(0.61 * 1 * data[1] * 16.02 * an * Math.Sqrt(2 * gc * Math.Abs(STORED_PRESSURE - ATMOSPHERIC_PRESSURE) / (data[1] * 16.02)) / (DAL_CAL.GET_TBL_3B21(1)), 2);
            }
            else
            {
                float R = 8.314f;
                k = (C_P() / (C_P() - R));
                p_trans = (float)(ATMOSPHERIC_PRESSURE * Math.Pow((k + 1) / 2, k / (k - 1)));
                if (STORED_PRESSURE > p_trans)
                {
                    float x = (float)((k * mw * gc / (R * STORED_TEMP)) * Math.Pow(2 / (k + 1), (k + 1) / (k - 1)));
                    W_max8 = (float)Math.Round(0.9 * an * STORED_PRESSURE * Math.Sqrt(Math.Abs(x)) / (DAL_CAL.GET_TBL_3B21(2)), 2);
                }
                else
                {
                    float x = (float)Math.Round((mw * gc / (R * STORED_TEMP)) * ((2 * k) / (k - 1)) * Math.Pow(ATMOSPHERIC_PRESSURE / STORED_PRESSURE, 2 / k) * (1 - Math.Pow(ATMOSPHERIC_PRESSURE / STORED_PRESSURE, (k - 1) / k)), 2);
                    W_max8 = (float)Math.Round(0.9 * an * STORED_PRESSURE * Math.Sqrt(Math.Abs(x)) / (DAL_CAL.GET_TBL_3B21(2)), 2);
                }
            }
            return W_max8;
        }
        private float mass_addn(int n)
        {
            float mass_addn = 0;
            float Wmax8 = W_max8();
            float Wn = W_n(n);
            mass_addn = 180 * Math.Min(Wmax8, Wn);
            return mass_addn;
        }
        public float mass_availn(int n)
        {
            float mass_availn = 0;
            float massaddn = mass_addn(n);
            mass_availn = Math.Min(MASS_COMPONENT + massaddn, MASS_INVERT);
            return mass_availn;
        }
        public float t_n(int n)
        {
            float t_n = 0;
            float Wn = W_n(n);
            t_n = (DAL_CAL.GET_TBL_3B21(3)) / Wn;
            return t_n;
        }
        public String releaseType(int n)
        {
            double tn = t_n(n);
            double dn = d_n(n);
            double massavailn = mass_availn(n);
            if (dn <= 6.35)
                return "Continuous";
            else if ((tn <= 180) || (massavailn > 4536))
                return "Instantaneous";
            else
                return "Continuous";
        }
        public float fact_di()
        {
            float fact_di = 0;
            if (DETECTION_TYPE == "A" && ISULATION_TYPE == "A")
            {
                fact_di = 0.25f;
            }
            else if (DETECTION_TYPE == "A" && ISULATION_TYPE == "B")
            {
                fact_di = 0.2f;
            }
            else if ((DETECTION_TYPE == "A" || DETECTION_TYPE == "B") && ISULATION_TYPE == "C")
            {
                fact_di = 0.1f;
            }
            else if (DETECTION_TYPE == "B" && ISULATION_TYPE == "B")
            {
                fact_di = 0.15f;
            }
            else fact_di = 0;
            return fact_di;
        }
        public float ld_n_max(int n)
        {
            float ld_max = 0;
            float dn = d_n(n);
            if (DETECTION_TYPE == "A" && ISULATION_TYPE == "A")
            {
                if (dn == 6.4f)
                    ld_max = 20;
                else if (dn == 25)
                    ld_max = 10;
                else if (dn == 102)
                    ld_max = 5;
                else
                    ld_max = 0;
            }
            else if (DETECTION_TYPE == "A" && ISULATION_TYPE == "B")
            {
                if (dn == 6.4f)
                    ld_max = 30;
                else if (dn == 25)
                    ld_max = 20;
                else if (dn == 102)
                    ld_max = 10;
                else
                    ld_max = 0;
            }
            else if (DETECTION_TYPE == "A" && ISULATION_TYPE == "C")
            {
                if (dn == 6.4f)
                    ld_max = 40;
                else if (dn == 25)
                    ld_max = 30;
                else if (dn == 102)
                    ld_max = 20;
                else
                    ld_max = 0;
            }
            else if ((ISULATION_TYPE == "A" || ISULATION_TYPE == "B") && DETECTION_TYPE == "B")
            {
                if (dn == 6.4f)
                    ld_max = 40;
                else if (dn == 25)
                    ld_max = 30;
                else if (dn == 102)
                    ld_max = 20;
                else
                    ld_max = 0;
            }
            else if (DETECTION_TYPE == "B" && ISULATION_TYPE == "C")
            {
                if (dn == 6.4f)
                    ld_max = 60;
                else if (dn == 25)
                    ld_max = 30;
                else if (dn == 102)
                    ld_max = 20;
                else
                    ld_max = 0;
            }
            else if (DETECTION_TYPE == "C" && (ISULATION_TYPE == "A" || ISULATION_TYPE == "B" || ISULATION_TYPE == "C"))
            {
                if (dn == 6.4f)
                    ld_max = 60;
                else if (dn == 25)
                    ld_max = 40;
                else if (dn == 102)
                    ld_max = 20;
                else
                    ld_max = 0;
            }
            else
                ld_max = 0;
            return ld_max;
        }
        public float rate_n(int n)
        {
            float rate_n = 0;
            float Wn = W_n(n);
            float factdi = fact_di();
            rate_n = Wn * (1 - factdi);
            if (rate_n == 0)
                return 1;
            else
                return rate_n;
        }
        public float ld_n(int n)
        {
            float ldmax = ld_n_max(n);
            if (ldmax != 0)
                return (float)Math.Min(mass_availn(n) / rate_n(n), 60 * ldmax);
            else
                return (float)(mass_availn(n) / rate_n(n));
        }
        public float mass_n(int n)
        {
            return Math.Min(rate_n(n) * ld_n(n), mass_availn(n));
        }
        public float fact_mit()
        {
            float fact_mit = 0;
            if (MITIGATION_SYSTEM == "Inventory blowdown, couple with isolation system classification B or higher")
                fact_mit = 0.25f;
            else if (MITIGATION_SYSTEM == "Fire water deluge system and monitors")
                fact_mit = 0.2f;
            else if (MITIGATION_SYSTEM == "Fire water monitors only")
                fact_mit = 0.05f;
            else
                fact_mit = 0.15f;
            return fact_mit;
        }
        public float eneff_n(int n)
        {
            float eff = (float)(4 * Math.Log10(DAL_CAL.GET_TBL_3B21(4) * mass_n(n)) - 15);
            if (eff == 0)
                return 1;
            else
                return eff;
        }
        public float a_cont(int select)
        {
            float[] data = DAL_CAL.GET_TBL_58(FLUID);
            float[] a_cont = { 0, 0, 0, 0 };
            if (GET_RELEASE_PHASE() == "Gas")
            {
                a_cont[0] = data[0];
                a_cont[1] = data[4];
                a_cont[2] = data[8];
                a_cont[3] = data[12];
            }
            else
            {
                a_cont[0] = data[2];
                a_cont[1] = data[6];
                a_cont[2] = data[10];
                a_cont[3] = data[14];
            }
            if (a_cont[select - 1] == 0)
                return 1;
            else
                return a_cont[select - 1];
        }
        public float b_cont(int select)
        {
            float[] data = DAL_CAL.GET_TBL_58(FLUID);
            float[] b_cont = { 0, 0, 0, 0 };
            if (GET_RELEASE_PHASE() == "Gas")
            {
                b_cont[0] = data[1];
                b_cont[1] = data[5];
                b_cont[2] = data[9];
                b_cont[3] = data[13];
            }
            else
            {
                b_cont[0] = data[3];
                b_cont[1] = data[7];
                b_cont[2] = data[11];
                b_cont[3] = data[15];
            }
            return b_cont[select - 1];
        }
        private float a_inj(int select)
        {
            float[] data = DAL_CAL.GET_TBL_59(FLUID);
            float[] a_inj = { 0, 0, 0, 0 };
            if (GET_RELEASE_PHASE() == "Gas")
            {
                a_inj[0] = data[0];
                a_inj[1] = data[4];
                a_inj[2] = data[8];
                a_inj[3] = data[12];
            }
            else
            {
                a_inj[0] = data[2];
                a_inj[1] = data[6];
                a_inj[2] = data[10];
                a_inj[3] = data[14];
            }
            return a_inj[select - 1];
        }
        private float b_inj(int select)
        {
            float[] data = DAL_CAL.GET_TBL_59(FLUID);
            float[] b_inj = { 0, 0, 0, 0 };//{ data[1], data[3], data[5], data[7], data[9], data[11], data[13], data[15] };
            if (GET_RELEASE_PHASE() == "Gas")
            {
                b_inj[0] = data[1];
                b_inj[1] = data[5];
                b_inj[2] = data[9];
                b_inj[3] = data[13];
            }
            else
            {
                b_inj[0] = data[3];
                b_inj[1] = data[7];
                b_inj[2] = data[11];
                b_inj[3] = data[15];
            }
            return b_inj[select - 1];
        }
        private float ca_cmdn_cont(int select, int n)
        {
            float ca_cmdn_cont = 0;
            String API_FLUID_TYPE = TYPE_FLUID();
            if ((GET_RELEASE_PHASE() == "Liquid") && (API_FLUID_TYPE == "TYPE 0"))
                ca_cmdn_cont = (float)Math.Round(Math.Min(a_cont(select) * Math.Pow(rate_n(n), b_cont(select)), DAL_CAL.GET_TBL_3B21(7)) * (1 - fact_mit()), 2);
            else
                ca_cmdn_cont = (float)Math.Round(a_cont(select) * Math.Pow(rate_n(n), b_cont(select)) * (1 - fact_mit()), 2);
            return ca_cmdn_cont;
        }
        private float effrate_n(int select, int n)
        {
            float effrate_n = 0;
            String API_FLUID_TYPE = TYPE_FLUID();
            if ((GET_RELEASE_PHASE() == "Liquid") && (API_FLUID_TYPE == "TYPE 0"))
                effrate_n = (float)Math.Round((1 / (DAL_CAL.GET_TBL_3B21(4)) * Math.Exp(Math.Log10(ca_cmdn_cont(select, n) / (a_cont(select) * (DAL_CAL.GET_TBL_3B21(8)))) * Math.Pow(b_cont(select), -1))), 2);
            else
                effrate_n = rate_n(n);
            return effrate_n;
        }
        private float ca_cmdn_inst(int select, int n)
        {
            float ca_cmdn_inst = 0;
            String API_FLUID_TYPE = TYPE_FLUID();
            if ((GET_RELEASE_PHASE() == "Liquid") && (API_FLUID_TYPE == "TYPE 0"))
                ca_cmdn_inst = (float)Math.Round(Math.Min(a_cont(select) * Math.Pow(mass_n(n), b_cont(select)), (DAL_CAL.GET_TBL_3B21(7))) * ((1 - fact_mit()) / eneff_n(n)), 2);
            else
                ca_cmdn_inst = (float)Math.Round(a_cont(select) * Math.Pow(mass_n(n), b_cont(select)) * (1 - fact_mit()), 2);
            return ca_cmdn_inst;
        }
        public float ca_injn_cont(int select, int n)
        {
            float ca_injn_cont;
            ca_injn_cont = (float)Math.Round(a_inj(select) * Math.Pow(effrate_n(select, n), b_inj(select)) * (1 - fact_mit()), 2);
            return ca_injn_cont;
        }
        public float ca_injn_inst(int select, int n)
        {
            float ca_injn_inst;
            ca_injn_inst = (float)Math.Round(a_inj(select) * Math.Pow(effrate_n(select, n), b_inj(select)) * ((1 - fact_mit()) / eneff_n(n)), 2);
            return ca_injn_inst;
        }
        public double fact_n_ic(int n)
        {
            String releasetype = releaseType(n);
            if (releasetype == "Continuous")
                return Math.Min(rate_n(n) / (DAL_CAL.GET_TBL_3B21(5)), 1.0);
            else
                return 1.0;
        }
        public float fact_ait()
        {
            float[] data = DAL_CAL.GET_TBL_52(FLUID);
            float fact_ait = 0;
            float ait = 273 + (float)Math.Round((data[9] - 32) / 1.8, 2);
            if ((STORED_TEMP + (DAL_CAL.GET_TBL_3B21(6))) <= ait)
                fact_ait = 0;
            else if ((STORED_TEMP - (DAL_CAL.GET_TBL_3B21(6))) >= ait)
                fact_ait = 1;
            else
                fact_ait = (STORED_TEMP - ait + (DAL_CAL.GET_TBL_3B21(6))) / (2 * (DAL_CAL.GET_TBL_3B21(6)));
            return fact_ait;
        }
        public float ca_cmdn_flame(int n)
        {
            float ca_cmdn_flame = 0;
            float caailcmdn = (float)(ca_cmdn_cont(2, n) * fact_n_ic(n) + ca_cmdn_inst(4, n) * (1 - fact_n_ic(n)));
            float caainlcmdn = (float)(ca_cmdn_cont(1, n) * fact_n_ic(n) + ca_cmdn_inst(3, n) * (1 - fact_n_ic(n)));
            ca_cmdn_flame = caailcmdn * fact_ait() + caainlcmdn * (1 - fact_ait());
            return ca_cmdn_flame;
        }
        public float ca_injn_flame(int n)
        {
            float caailinjn = (float)(Math.Abs(ca_injn_cont(2, n)) * fact_n_ic(n) + Math.Abs(ca_injn_inst(4, n)) * (1 - fact_n_ic(n)));
            float caainlinjn = (float)(Math.Abs(ca_injn_cont(1, n)) * fact_n_ic(n) + Math.Abs(ca_injn_inst(3, n)) * (1 - fact_n_ic(n)));
            return caailinjn * fact_ait() + caainlinjn * (1 - fact_ait());

        }
        private Boolean checkFlame()
        {
            Boolean check = false;
            string[] itemsFrammable = { "C1-C2", "C3-C4", "C5", "C6-C8", "C9-C12", "C13-C16", "C17-C25", "C25+", "H2", "H2S", "HF", "CO", "DEE", "Methanol", "PO", "Styrene", "Aromatics", "EEA", "EE", "EG", "EO" };
            for (int i = 0; i < itemsFrammable.Length; i++)
            {
                if (FLUID == itemsFrammable[i])
                {
                    check = true;
                    break;
                }
            }
            return check;
        }
        public float ca_cmd_flame()
        {
            if (!checkFlame())
            {
                return 0;
            }
            else
            {
                float t = 0;
                API_COMPONENT_TYPE obj = GET_DATA_API_COM();
                t = obj.GFFSmall * ca_cmdn_flame(1) + obj.GFFMedium * ca_cmdn_flame(2) + obj.GFFLarge * ca_cmdn_flame(3) + obj.GFFRupture * ca_cmdn_flame(4);
                float ca_cmd_flame = t / obj.GFFTotal;
                return Math.Abs(ca_cmd_flame);
            }

        }
        public float ca_inj_flame()
        {
            if (!checkFlame())
            {
                return 0;
            }
            else
            {
                float t;
                API_COMPONENT_TYPE obj = GET_DATA_API_COM();
                t = obj.GFFSmall * ca_injn_flame(1) + obj.GFFMedium * ca_injn_flame(2) + obj.GFFLarge * ca_injn_flame(3) + obj.GFFRupture * ca_injn_flame(4);
                float ca_inj = t / obj.GFFTotal;
                return Math.Abs(ca_inj);
            }
        }
        // TOXIC CONSEQUENCE
        public float rate_tox_n(int n)
        {
            return TOXIC_PERCENT * W_n(n) / 100;
        }
        public float mass_tox_n(int n)
        {
            return TOXIC_PERCENT * mass_n(n) / 100;
        }
        public TOXIC_511_512 getToxic(string type)
        {
            TOXIC_511_512 obj = new TOXIC_511_512();
            List<TOXIC_511_512> list = DAL_CAL.GET_TBL_511_512();
            if (type == "Instantaneous")
            {
                RELEASE_DURATION = "Instantaneous Releases";
            }
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i].ToxicName == FLUID && list[i].ReleaseDuration == RELEASE_DURATION)
                    obj = list[i];
            }
            return obj;
        }
        public TOXIC_513 getToxic513(string type)
        {
            TOXIC_513 obj = new TOXIC_513();
            List<TOXIC_513> list = DAL_CAL.GET_TBL_513();
            if (FLUID_PHASE == "Vapor")
                TOXIC_PHASE = "Gas";
            else
                TOXIC_PHASE = "Liquid";
            if (type == "Instantaneous")
            {
                RELEASE_DURATION = "Instantaneous Releases";
            }
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i].TOXIC_NAME == FLUID && list[i].TOXIC_TYPE == TOXIC_PHASE && list[i].DURATION == RELEASE_DURATION)
                    obj = list[i];
            }
            return obj;
        }
        public float ca_injn_tox(int n)
        {
            float C8 = DAL_CAL.GET_TBL_3B21(8);
            float C4 = DAL_CAL.GET_TBL_3B21(4);
            String releasetype = releaseType(n);
            TOXIC_511_512 obj = getToxic(releasetype);
            TOXIC_513 obj1 = getToxic513(releasetype);
            if (obj.ToxicName == "HF" || obj.ToxicName == "H2S")
            {
                double log = 0;
                if (releasetype == "Continuous")
                    log = obj.a * Math.Log10(C4 * rate_tox_n(n)) + obj.b;
                else
                    log = obj.a * Math.Log10(C4 * mass_tox_n(n)) + obj.b;
                return (float)Math.Round(C8 * Math.Pow(10, log), 4);
            }
            else if (obj.ToxicName == "Ammonia" || obj.ToxicName == "Chlorine")
            {
                if (releasetype == "Continuous")
                    return (float)Math.Round(obj.a * Math.Pow(rate_tox_n(n), obj.b), 4);
                else
                    return (float)Math.Round(obj.a * Math.Pow(mass_tox_n(n), obj.b), 4);
            }
            else if (obj1.TOXIC_NAME == "AlCl3" || obj1.TOXIC_NAME == "CO" || obj1.TOXIC_NAME == "HCl" || obj1.TOXIC_NAME == "Nitric Acid" || obj1.TOXIC_NAME == "NO2"
                    || obj1.TOXIC_NAME == "Phosgene" || obj1.TOXIC_NAME == "TDI" || obj1.TOXIC_NAME == "EE" || obj1.TOXIC_NAME == "EO" || obj1.TOXIC_NAME == "PO")
            {
                if (releasetype == "Continuous")
                    return (float)Math.Round(obj1.a * Math.Pow(rate_tox_n(n), obj1.b), 4);
                else
                    return (float)Math.Round(obj1.a * Math.Pow(mass_tox_n(n), obj1.b), 4);
            }
            else
            {
                return 0;
            }
        }
        private Boolean checkToxic()
        {
            Boolean check = false;
            string[] itemsToxic = { "Nitric Acid", "AlCl3", "TDI", "EE" };
            if (FLUID == "H2S" && TOXIC_PERCENT > (100 * Math.Pow(10, -4)))
                check = true;
            else if (FLUID == "HF" && TOXIC_PERCENT > (30 * Math.Pow(10, -4)))
                check = true;
            else if (FLUID == "CO" && TOXIC_PERCENT > (1200 * Math.Pow(10, -4)))
                check = true;
            else if (FLUID == "HCl" && TOXIC_PERCENT > (50 * Math.Pow(10, -4)))
                check = true;
            else if (FLUID == "NO2" && TOXIC_PERCENT > (20 * Math.Pow(10, -4)))
                check = true;
            else if (FLUID == "Phosgene" && TOXIC_PERCENT > (2 * Math.Pow(10, -4)))
                check = true;
            else if (FLUID == "PO" && TOXIC_PERCENT > (400 * Math.Pow(10, -4)))
                check = true;
            else if (FLUID == "EO" && TOXIC_PERCENT > (800 * Math.Pow(10, -4)))
                check = true;
            else if (FLUID == "Ammonia" && TOXIC_PERCENT > (10 * Math.Pow(10, -4)))
                check = true;
            else if (FLUID == "Chlorine" && TOXIC_PERCENT > (10 * Math.Pow(10, -4)))
                check = true;
            else
            {
                for (int i = 0; i < itemsToxic.Length; i++)
                {
                    if (FLUID == itemsToxic[i])
                    {
                        check = true;
                        break;
                    }
                }
            }
            return check;
        }
        public float ca_inj_tox()
        {
            API_COMPONENT_TYPE obj = GET_DATA_API_COM();
            if (!checkToxic())
            {
                return 0;
            }
            else
            {
                float t = 0;
                t = obj.GFFSmall * ca_injn_tox(1) + obj.GFFMedium * ca_injn_tox(2) + obj.GFFLarge * ca_injn_tox(3) + obj.GFFRupture * ca_injn_tox(4);
                float ca_inj_tox = t / obj.GFFTotal;
                return Math.Abs(ca_inj_tox);
            }

        }
        // Step 10 non flammable non toxic consequence
        public float ca_injn_contnfnt(int n)
        {
            float ca_injn_cont = 0;
            float g = (float)Math.Round(2696 - 21.9 * (DAL_CAL.GET_TBL_3B21(11)) * (STORED_PRESSURE - ATMOSPHERIC_PRESSURE) + 1.474 * Math.Pow(((DAL_CAL.GET_TBL_3B21(11)) * (STORED_PRESSURE - ATMOSPHERIC_PRESSURE)), 2), 4);
            float h = (float)Math.Round(0.31 - 0.00032 * Math.Pow(((DAL_CAL.GET_TBL_3B21(11)) * (STORED_PRESSURE - ATMOSPHERIC_PRESSURE) - 40), 2), 4);
            if (FLUID == "Steam")
                ca_injn_cont = (float)Math.Round((DAL_CAL.GET_TBL_3B21(9)) * rate_n(n), 4);
            else
                ca_injn_cont = (float)Math.Round(0.2 * (DAL_CAL.GET_TBL_3B21(8)) * g * Math.Pow((DAL_CAL.GET_TBL_3B21(4)) * rate_n(n), h), 2);
            return ca_injn_cont;
        }
        public float ca_injn_instnfnt(int n)
        {
            float ca_injn_inst = 0;
            if (FLUID == "Steam")
                ca_injn_inst = (float)Math.Round((DAL_CAL.GET_TBL_3B21(10)) * Math.Pow(mass_n(n), 0.6384), 4);
            else
                ca_injn_inst = 0;
            return ca_injn_inst;
        }
        public float fact_n_icnfnt(int n)
        {
            float fact_n_icnfnt = 0;
            if (FLUID == "Steam")
                fact_n_icnfnt = Math.Min(rate_n(n) / (DAL_CAL.GET_TBL_3B21(5)), 1);
            else
                fact_n_icnfnt = 0;
            return fact_n_icnfnt;
        }
        public float ca_injn_leaknfnt(int n)
        {
            return ca_injn_instnfnt(n) * fact_n_icnfnt(n) + ca_injn_contnfnt(n) * (1 - fact_n_icnfnt(n));
        }
        private Boolean checkNone()
        {
            Boolean check = false;
            string[] itemsNoneTF = { "Steam", "Acid", "Caustic" };
            for (int i = 0; i < itemsNoneTF.Length; i++)
            {
                if (FLUID == itemsNoneTF[i])
                {
                    check = true;
                    break;
                }
            }
            return check;
        }
        public float ca_inj_nfnt()
        {
            if (!checkNone())
            {
                return 0;
            }
            else
            {
                float t = 0;
                API_COMPONENT_TYPE obj = GET_DATA_API_COM();
                t = obj.GFFSmall * ca_injn_leaknfnt(1) + obj.GFFMedium * ca_injn_leaknfnt(2) + obj.GFFLarge * ca_injn_leaknfnt(3) + obj.GFFRupture * ca_injn_leaknfnt(4);
                float ca_inj_nfnt = t / obj.GFFTotal;
                return Math.Abs(ca_inj_nfnt);
            }
        }
        // Step 11 Component and Injure consequence
        public float ca_cmd()
        {
            return ca_cmd_flame();
        }
        public float ca_inj()
        {
            float cainjflame = ca_inj_flame();
            float cainjtox = ca_inj_tox();
            float cainjnfnt = ca_inj_nfnt();
            return Math.Max(Math.Max(cainjflame, cainjtox), cainjnfnt);
        }
        // Step 12: financial
        public float fc_cmd()
        {
            float fc_cmd = 0;
            API_COMPONENT_TYPE obj = GET_DATA_API_COM();
            float t = 0;
            t = obj.GFFSmall * obj.HoleCostSmall + obj.GFFMedium * obj.HoleCostMedium + obj.GFFLarge * obj.HoleCostLarge + obj.GFFRupture * obj.HoleCostRupture;
            fc_cmd = t * MATERIAL_COST / obj.GFFTotal;
            return fc_cmd;
        }
        public float fc_affa()
        {
            float fc_affa = 0;
            float cacmd = ca_cmd();
            fc_affa = cacmd * EQUIPMENT_COST;
            return fc_affa;
        }
        public float outage_cmd()
        {
            float outage_cmd = 0;
            float t = 0;
            API_COMPONENT_TYPE obj = GET_DATA_API_COM();
            t = obj.GFFSmall * obj.OutageSmall + obj.GFFMedium * obj.OutageMedium + obj.GFFLarge * obj.OutageLarge + obj.GFFRupture * obj.GFFRupture;
            outage_cmd = t / obj.GFFTotal;
            return outage_cmd;
        }
        public float outage_affa()
        {
            float outage_affa = 0;
            float fcaffa = Math.Abs(fc_affa());
            float b = (float)(1.242 + 0.585 * Math.Log10(fcaffa * Math.Pow(10, -6)));
            outage_affa = (float)Math.Pow(10, b);
            return outage_affa;
        }
        public float fc_prod()
        {
            float fc_prod = (outage_cmd() + outage_affa()) * PRODUCTION_COST;
            return fc_prod;
        }
        public float fc_inj()
        {
            float fc_inj = ca_inj() * PERSON_DENSITY * INJURE_COST;
            return fc_inj;
        }
        public float vol_n_env(int n)
        {
            float vol_n_env = 0;
            float massn = mass_n(n);
            float frac_evap = 1;
            switch (FLUID)
            {
                case "C6-C8":
                case "Acid":
                    frac_evap = 0.9f;
                    break;
                case "C9-C12":
                    frac_evap = 0.5f;
                    break;
                case "C13-C16":
                    frac_evap = 0.1f;
                    break;
                case "C17-C25":
                    frac_evap = 0.05f;
                    break;
                case "C25+":
                    frac_evap = 0.02f;
                    break;
                case "Nitric Acid":
                    frac_evap = 0.8f;
                    break;
                case "NO2":
                case "EE":
                    frac_evap = 0.75f;
                    break;
                case "TDI":
                    frac_evap = 0.15f;
                    break;
                case "Styrene":
                    frac_evap = 0.6f;
                    break;
                case "EEA":
                    frac_evap = 0.65f;
                    break;
                case "EG":
                    frac_evap = 0.45f;
                    break;
                default:
                    frac_evap = 1;
                    break;
            }
            float[] data = DAL_CAL.GET_TBL_52(FLUID);
            vol_n_env = (float)((DAL_CAL.GET_TBL_3B21(13)) * massn * (1 - frac_evap) / (data[1] * 16.02));
            return vol_n_env;
        }
        public float fc_environ()
        {
            API_COMPONENT_TYPE obj = GET_DATA_API_COM();
            float fc_environ = 0;
            float t = 0;
            t = obj.GFFSmall * vol_n_env(1) + obj.GFFMedium * vol_n_env(2) + obj.GFFLarge * vol_n_env(3) + obj.GFFRupture * vol_n_env(4);
            fc_environ = t * ENVIRON_COST / obj.GFFTotal;
            return fc_environ;
        }
        public float fc()
        {
            float fc = 0;
            float fccmd = fc_cmd();
            float fcaffa = fc_affa();
            float fcprod = fc_prod();
            float fcinj = fc_inj();
            float fcenviron = fc_environ();
            fc = fccmd + fcaffa + fcprod + fcinj + fcenviron;
            return fc;
        }

        // Storage tank
        /*
         * Tank Shell Course
         * */
        public float FLUID_HEIGHT { set; get; }
        public float Cd = 0.61f;
        public float SHELL_COURSE_HEIGHT { set; get; }
        public float TANK_DIAMETER { set; get; }
        public Boolean PREVENTION_BARRIER { set; get; }//Release Prevention Barrier
        public String EnvironSensitivity { set; get; }
        public float P_lvdike { set; get; }
        public float P_onsite { set; get; }
        public float P_offsite { set; get; }

        public float W_n_Tank(int n)
        {
            float Wn_tank = (float)(DAL_CAL.GET_TBL_3B21(32) * Cd * a_n(n) * Math.Sqrt(2 * FLUID_HEIGHT));
            return Wn_tank > 0 ? Wn_tank : 1;
        }

        public float BBL_TOTAL_SHELL()
        {
            return (float)(Math.PI * Math.Pow(TANK_DIAMETER, 2)) * FLUID_HEIGHT / (4 * (DAL_CAL.GET_TBL_3B21(13)));
        }
        private float Bbl_avail(int n)
        {
            int i = (int)(FLUID_HEIGHT / SHELL_COURSE_HEIGHT);
            //if (n <= i)
            //{
            //    return (float)(Math.PI * Math.Pow(TANK_DIAMETER, 2)) * (FLUID_HEIGHT - (n - 1) * SHELL_COURSE_HEIGHT) / (4 * (DAL_CAL.GET_TBL_3B21(13)));
            //}
            //else
            //{
            //    return 0;
            //}
            float Bbl_avail =  (float)(Math.PI * Math.Pow(TANK_DIAMETER, 2)) * (FLUID_HEIGHT - (n - 1) * SHELL_COURSE_HEIGHT) / (4 * (DAL_CAL.GET_TBL_3B21(13)));
            return Bbl_avail > 0 ? Bbl_avail : 0;
        }

        public float ld_tank(int n)
        {
            if (d_n(n) <= 3.17)
            {
                float ld_tank_shell =  Math.Min(Bbl_avail(n) / W_n_Tank(n), 7);
                return ld_tank_shell > 0 ? ld_tank_shell : 1;
            }
            else
            {
                return 1;
            }
        }
        public float Bbl_leak_n(int n)
        {
            //if (d_n(n) <= 50.8)
            //{
                return Math.Min(W_n_Tank(n) * ld_tank(n), Bbl_avail(n));
            //}
            //else
                //return Bbl_avail(n);
        }
        public int[] getCost()
        {
            int[] costTANK = { 0, 0, 0, 0, 0, 0 };
            if (EnvironSensitivity == "High")
            {
                costTANK[0] = 10;
                costTANK[1] = 50;
                costTANK[2] = 500;
                costTANK[3] = 3000;
                costTANK[4] = 10000;
                costTANK[5] = 5000;
            }
            else if (EnvironSensitivity == "Medium")
            {
                costTANK[0] = 10;
                costTANK[1] = 50;
                costTANK[2] = 250;
                costTANK[3] = 1500;
                costTANK[4] = 5000;
                costTANK[5] = 1500;
            }
            else if (EnvironSensitivity == "Low")
            {
                costTANK[0] = 10;
                costTANK[1] = 50;
                costTANK[2] = 100;
                costTANK[3] = 500;
                costTANK[4] = 1000;
                costTANK[5] = 500;
            }
            else
            {
                costTANK[0] = 0;
                costTANK[1] = 0;
                costTANK[2] = 0;
                costTANK[3] = 0;
                costTANK[4] = 0;
                costTANK[5] = 0;
            }
            return costTANK;
        }
        public float Bbl_leak_release()
        {
            float sum = 0;
            API_COMPONENT_TYPE obj = GET_DATA_API_COM();
            sum = Bbl_leak_n(1) * obj.GFFSmall + Bbl_leak_n(2) * obj.GFFMedium + Bbl_leak_n(3) * obj.GFFLarge;
            return sum / obj.GFFTotal;
        }
        public float Bbl_leak_indike()
        {
            return Bbl_leak_release() * (1 - P_lvdike / 100);
        }
        public float Bbl_leak_ssonsite()
        {
            return P_onsite * (Bbl_leak_release() - Bbl_leak_indike()) / 100;
        }
        public float Bbl_leak_ssoffsite()
        {
            return P_offsite * (Bbl_leak_release() - Bbl_leak_indike() - Bbl_leak_ssonsite()) / 100;
        }
        public float Bbl_leak_water()
        {
            return Bbl_leak_release() - (Bbl_leak_indike() + Bbl_leak_ssonsite() + Bbl_leak_ssoffsite());
        }
        //Step 5: Compute FC_leakage_environ
        public float FC_leak_environ()
        {
            int[] cost = getCost();
            float fc_leak_env =  Bbl_leak_indike() * cost[0] + Bbl_leak_ssonsite() * cost[1] + Bbl_leak_ssoffsite() * cost[2] + Bbl_leak_water() * cost[5];
            return fc_leak_env > 0 ? fc_leak_env : 0;
        }
        //Step 6: Determine Bbl_rupture_release
        public float Bbl_rupture_release()
        {
            API_COMPONENT_TYPE obj = GET_DATA_API_COM();
            return Bbl_leak_n(4) * obj.GFFRupture / obj.GFFTotal;
        }
        //Step 7: Compute 4 parameter Bbl_rupture_indike, Bbl_rupture_ssonsite, Bbl_rupture_ssoffsite, Bbl_rupture_water
        public float Bbl_rupture_indike()
        {
            return Bbl_rupture_release() * (1 - P_lvdike / 100);
        }
        public float Bbl_rupture_ssonsite()
        {
            return P_onsite * (Bbl_rupture_release() - Bbl_rupture_indike()) / 100;
        }
        public float Bbl_rupture_ssoffsite()
        {
            return P_offsite * (Bbl_rupture_release() - Bbl_rupture_indike() - Bbl_rupture_ssonsite()) / 100;
        }
        public float Bbl_rupture_water()
        {
            return Bbl_rupture_release() - (Bbl_rupture_indike() + Bbl_rupture_ssonsite() + Bbl_rupture_ssoffsite());
        }
        //Step 8: Compute FC_rupture_environ
        public float FC_rupture_environ()
        {
            int[] cost = getCost();
            float fc_rupture_env =  Bbl_rupture_indike() * cost[0] + Bbl_rupture_ssonsite() * cost[1] + Bbl_rupture_ssoffsite() * cost[2] + Bbl_rupture_water() * cost[5];
            return fc_rupture_env > 0 ? fc_rupture_env : 0;
        }
        //Step 9: Compute FC_environ
        public float FC_environ_shell()
        {
            return FC_leak_environ() + FC_rupture_environ();
        }
        //Step 12: Compute total Financial
        public float FC_PROD_SHELL()
        {
            API_COMPONENT_TYPE obj = GET_DATA_API_COM();
            float t = obj.GFFSmall * obj.OutageSmall + obj.GFFMedium * obj.OutageMedium + obj.GFFLarge * obj.OutageLarge + obj.GFFRupture * obj.OutageRupture;
            return t*PRODUCTION_COST / obj.GFFTotal;
        }
        public double FC_total_shell()
        {
            return fc_cmd() + FC_environ_shell() + FC_PROD_SHELL();
        }

        /*
         * TANK BOTTOM
         * */
        public String Soil_type { set; get; }
        public String TANK_FLUID { set; get; }
        public float Swg { set; get; } //total distance to the ground water underneath the tank
        private float uw = 0.001f; //tra bang tren internet tai 27C
        private float pw = 1000; // kg/m3

        //STEP 1: Determine Release Rate and Volum
        private int n_rh()
        {
            return (int)Math.Round(Math.Max(Math.Pow(TANK_DIAMETER / DAL_CAL.GET_TBL_3B21(36), 2), 1), 0);
        }
        private float[] k_h_bottom()
        {
            float[] k_h = { 0, 0, 0 };
            if (Soil_type == "Coarse Sand")
            {
                k_h[0] = 0.1f;
                k_h[1] = 0.01f;
                k_h[2] = 0.33f;
            }
            else if (Soil_type == "Fine Sand")
            {
                k_h[0] = 0.01f;
                k_h[1] = 0.001f;
                k_h[2] = 0.33f;
            }
            else if (Soil_type == "Very Fine Sand")
            {
                k_h[0] = (float)Math.Pow(10, -3);
                k_h[1] = (float)Math.Pow(10, -5);
                k_h[2] = 0.33f;
            }
            else if (Soil_type == "Silt")
            {
                k_h[0] = (float)Math.Pow(10, -5);
                k_h[1] = (float)Math.Pow(10, -6);
                k_h[2] = 0.41f;
            }
            else if (Soil_type == "Sandy Clay")
            {
                k_h[0] = (float)Math.Pow(10, -6);
                k_h[1] = (float)Math.Pow(10, -7);
                k_h[2] = 0.45f;
            }
            else if (Soil_type == "Clay")
            {
                k_h[0] = (float)Math.Pow(10, -7);
                k_h[1] = (float)Math.Pow(10, -8);
                k_h[2] = 0.5f;
            }
            else
            {
                k_h[0] = (float)Math.Pow(10, -10);
                k_h[1] = (float)Math.Pow(10, -11);
                k_h[2] = 0.99f;
            }
            return k_h;
        }
        public float k_h_water()
        {
            float[] k_h = k_h_bottom();
            return (DAL_CAL.GET_TBL_3B21(31)) * (k_h[0] + k_h[1]) / 2;
        }
        // W_n tank bottom cho d1 va d4
        public float rate_n_tank_bottom(int n)
        {
            if (k_h_water() > (DAL_CAL.GET_TBL_3B21(34)) * Math.Pow(d_n(n), 2))
                return (float)(DAL_CAL.GET_TBL_3B21(33) * Math.PI * d_n(n) * Math.Sqrt(2 * 1 * FLUID_HEIGHT) * n_rh());
            else
                return (float)(DAL_CAL.GET_TBL_3B21(35) * 0.21 * Math.Pow(d_n(n), 0.2) * Math.Pow(FLUID_HEIGHT, 0.9) * Math.Pow(k_h_water(), 0.74) * n_rh());
        }
        public float t_ld_tank_bottom()
        {
            if (Soil_type == "Concrete-Asphalt") return 7;
            else if (PREVENTION_BARRIER) return 30;
            else return 360;
        }

        public float BBL_TOTAL_TANKBOTTOM()
        {
            return (float)(Math.PI * Math.Pow(TANK_DIAMETER, 2) * FLUID_HEIGHT) / (4 * (DAL_CAL.GET_TBL_3B21(13)));
        }
        // ld_n cho truong hop d1 va d4
        public float ld_n_tank_bottom(int n)
        {
            float Bbl_total_tank_bottom = (float)(Math.PI * Math.Pow(TANK_DIAMETER, 2) * FLUID_HEIGHT) / (4 * (DAL_CAL.GET_TBL_3B21(13)));
            return Math.Min(Bbl_total_tank_bottom / rate_n_tank_bottom(n), t_ld_tank_bottom());
        }
        public float Bbl_leak_n_bottom(int n)
        {
            float Bbl_total_tank_bottom = (float)(Math.PI * Math.Pow(TANK_DIAMETER, 2) * FLUID_HEIGHT) / (4 * (DAL_CAL.GET_TBL_3B21(13)));
            return Math.Min(rate_n_tank_bottom(n) * ld_n_tank_bottom(n), Bbl_total_tank_bottom);
        }
        //STEP 2: Consequence Tank Bottom
        private float[] GET_PL_UL()
        {
            float[] data = { 0, 0 };
            if (TANK_FLUID == "Gasoline")
            {
                data[0] = 684.018f;
                data[1] = (float)(4.01 * Math.Pow(10, -3));
            }
            else if (TANK_FLUID == "Light Diesel Oil")
            {
                data[0] = 734.011f;
                data[1] = (float)(1.04 * Math.Pow(10, -3));
            }
            else if (TANK_FLUID == "Heavy Diesel Oil")
            {
                data[0] = 764.527f;
                data[1] = (float)(2.46 * Math.Pow(10, -3));
            }
            else if (TANK_FLUID == "Fuel Oil")
            {
                data[0] = 775.019f;
                data[1] = (float)(3.69 * Math.Pow(10, -2));
            }
            else if (TANK_FLUID == "Crude Oil")
            {
                data[0] = 775.019f;
                data[1] = (float)(3.69 * Math.Pow(10, -2));
            }
            else
            {
                data[0] = 900.026f;
                data[1] = (float)(4.6 * Math.Pow(10, -2));
            }
            return data;
        }
        public float k_h_prod()
        {
            float[] pl_ul = GET_PL_UL();
            return k_h_water() * (pl_ul[0] / pw) * (uw / pl_ul[1]);
        }
        public float vel_s_prod()
        {
            float[] kh = k_h_bottom();
            return k_h_prod() / kh[2];
        }
        public float t_gl_bottom()
        {
            return Swg / vel_s_prod();
        }
        public float Bbl_leak_groundwater(int n)
        {
            if (t_gl_bottom() > t_ld_tank_bottom())
                return Bbl_leak_n_bottom(n) * ((t_gl_bottom() - t_ld_tank_bottom()) / t_gl_bottom());
            else
                return 0;
        }
        public float Bbl_leak_subsoil(int n)
        {
            return Bbl_leak_n_bottom(n) - Bbl_leak_groundwater(n);
        }
        public float FC_leak_environ_bottom()
        {
            float sum = 0;
            int[] cost = getCost();
            API_COMPONENT_TYPE obj = GET_DATA_API_COM();
            sum = Bbl_leak_groundwater(1) * cost[4] + Bbl_leak_subsoil(1) * cost[3];
            return sum * obj.GFFSmall / obj.GFFTotal;
        }
        public float Bbl_rupture_release_bottom()
        {
            API_COMPONENT_TYPE obj = GET_DATA_API_COM();
            float Bbl_total_tank_bottom = (float)(Math.PI * Math.Pow(TANK_DIAMETER, 2) * FLUID_HEIGHT) / (4 * (DAL_CAL.GET_TBL_3B21(13)));
            return (Bbl_total_tank_bottom * obj.GFFRupture) / obj.GFFTotal;
        }
        public float Bbl_rupture_indike_bottom()
        {
            float indike =  Bbl_rupture_release_bottom() * (1 - P_lvdike / 100);
            return indike > 0 ? indike : 0;
        }
        public float Bbl_rupture_ssonsite_bottom()
        {
            float onsite = P_onsite * (Bbl_rupture_release_bottom() - Bbl_rupture_indike_bottom()) / 100;
            return onsite > 0 ? onsite : 0;
        }
        public float Bbl_rupture_ssoffsite_bottom()
        {
            float offsite =  P_offsite * (Bbl_rupture_release_bottom() - Bbl_rupture_indike_bottom() - Bbl_rupture_ssonsite_bottom()) / 100;
            return offsite > 0 ? offsite : 0;
        }
        public float Bbl_rupture_water_bottom()
        {
            float water =  Bbl_rupture_release_bottom() - (Bbl_rupture_indike_bottom() + Bbl_rupture_ssonsite_bottom() + Bbl_rupture_ssoffsite_bottom());
            return water > 0 ? water : 0;
        }
        public float FC_rupture_environ_bottom()
        {
            int[] cost = getCost();
            return Bbl_rupture_indike_bottom() * cost[0] + Bbl_rupture_ssonsite_bottom() * cost[1] + Bbl_rupture_ssoffsite_bottom() * cost[2] + Bbl_rupture_water_bottom() * cost[5];
        }
        public float FC_environ_bottom()
        {
            return FC_leak_environ_bottom() + FC_rupture_environ_bottom();
        }
        public float FC_cmd_bottom()
        {
            float sum = 0;
            API_COMPONENT_TYPE obj = GET_DATA_API_COM();
            sum = (float)(obj.GFFSmall * obj.HoleCostSmall + obj.GFFMedium * obj.HoleCostMedium + obj.GFFLarge * obj.HoleCostLarge + obj.HoleCostRupture * Math.Pow(TANK_DIAMETER / DAL_CAL.GET_TBL_3B21(36), 2));
            return sum * MATERIAL_COST / obj.GFFTotal;
        }
        public double FC_total_bottom()
        {
            return FC_environ_bottom() + FC_cmd_bottom() + FC_PROD_SHELL();
        }

        
    }
    //class CALevel2
    //{
    //    API_COMPONENT_TYPE_BUS busAPI_COM = new API_COMPONENT_TYPE_BUS();
    //    MSSQL_RBI_CAL_ConnUtils DAL_CAL = new MSSQL_RBI_CAL_ConnUtils();
    //    // material
    //    public String material { set; get; }
    //    // fluid
    //    public String fluid { set; get; }
    //    // component type
    //    public String componentType { set; get; }
    //    // phase of stored fluid
    //    public String fluidPhase { set; get; }
    //    // poolfire type
    //    public String poolfireType { set; get; }
    //    // equipcost
    //    public float equipcost { set; get; }
    //    // mass fraction liquid
    //    public float frac_l { set; get; }
    //    // vapor density
    //    public float p_v { set; get; }
    //    // fraction of fluid flashed
    //    public float frac_fsh { set; get; }
    //    // liquid density
    //    public float p_l { set; get; }
    //    // stored pressure kPa
    //    public float p_s { set; get; }
    //    // atmospheric pressure kPa
    //    public float p_atm { set; get; }
    //    // stored fluid's saturation pressure kPa
    //    public float psat_s { set; get; }
    //    // stored temperature K
    //    public float t_s { set; get; }
    //    // bubble point temperature for flashed liquid K
    //    public float t_b { set; get; }
    //    // dew point temperature for flashed vapor K
    //    public float t_d { set; get; }
    //    // atmosphere temperature K
    //    public float t_atm { set; get; }
    //    // time for steady release of fluid
    //    public float t_pn { set; get; }
    //    // specific heat of pool liquid J/kg-K
    //    public float c_pl { set; get; }
    //    // reynold constant
    //    public float r_e { set; get; }
    //    // fluid mass
    //    public float mass_inv { set; get; }
    //    // component mass
    //    public float mass_comp { set; get; }
    //    // mass of flammable material in vapor cloud
    //    public float mass_vce { set; get; }
    //    // detection type
    //    public float detectionType { set; get; }
    //    // isolation type
    //    public float isolationType { set; get; }
    //    // mass fraction of release rate
    //    public float mfrac_flame { set; get; }
    //    // volume of liquid to establish fire pool m3
    //    public float V_pn { set; get; }
    //    // bubble-point pressure, corresponding to the ground temperature kPa
    //    public float P_bg { set; get; }
    //    // wind speed m/s
    //    public float u_w { set; get; }
    //    // area surface type
    //    public String surface { set; get; }
    //    // ground temperature K
    //    public float t_g { set; get; }
    //    // ambient condition
    //    public String ambientCondition { set; get; }
    //    // humidity %
    //    public float RH { set; get; }
    //    // mole fraction of release rate
    //    public float molefrac_tox { set; get; }
    //    // toxic Component
    //    public String toxicComponent { set; get; }
    //    // criteria
    //    public String criteria { set; get; }
    //    // grade level cloud
    //    public float ca_n_cloud { set; get; }
    //    //represent Fluid
    //    public String representFluid { set; get; }
    //    // moles flash from liquid to vapor
    //    public float n_v { set; get; }
    //    // outage multiplier
    //    public float outage_mult { set; get; }
    //    // production cost ($)
    //    public float prodcost { set; get; }
    //    // injury cost that company would be incurred($)
    //    public float injcost { set; get; }
    //    // environment cost
    //    public float envcost { set; get; }
    //    // population density (person/m^2)
    //    public float popdens { set; get; }
    //    // m hoang thieu tu day tro xuong
    //    public float xs_cmdn_pool { set; get; }
    //    //
    //    public float xs_injn_pool { set; get; }
    //    //
    //    public float xs_cmdn_jet { set; get; }
    //    //
    //    public float xs_injn_jet { set; get; }
    //    //
    //    public float xs_cmdn_fb { set; get; }
    //    //
    //    public float xs_injn_fb { set; get; }
    //    //
    //    public float xs_cmdn_vce { set; get; }
    //    //
    //    public float xs_injn_vce { set; get; }
    //    // grade cloud area
    //    public float gradecloudArea { set; get; }

    //    // 
    //    public float xs_cmdn_pexp { set; get; }
    //    // 
    //    public float xs_cmdn_bleve { set; get; }
    //    // 
    //    public float xs_injn_pexp { set; get; }
    //    // 
    //    public float xs_injn_bleve { set; get; }
    //    // Step 2 release rate calculation
    //    // 2.1
    //    public float d_n(int n)
    //    {
    //        float dn = 0;
    //        if (n == 1)
    //            dn = 0.25f;
    //        else if (n == 2)
    //            dn = 1;
    //        else if (n == 3)
    //            dn = 4;
    //        else
    //            dn = 16;
    //        return dn;
    //    }
    //    // 2.2
    //    public float gff_n(int n)
    //    {
    //        float gff_n = 0;
    //        if (n == 1)
    //        {
    //            gff_n = busAPI_COM.GetGFFn(componentType, "GFFSmall");
    //        }
    //        else if (n == 2)
    //        {
    //            gff_n = busAPI_COM.GetGFFn(componentType, "GFFMedium");
    //        }
    //        else if (n == 3)
    //        {
    //            gff_n = busAPI_COM.GetGFFn(componentType, "GFFLarge");
    //        }
    //        else
    //        {
    //            gff_n = busAPI_COM.GetGFFn(componentType, "GFFRupture");
    //        }
    //        return gff_n;

    //    }

    //    public float gff_total()
    //    {
    //        float gff_total = 1;
    //        gff_total = busAPI_COM.getGFFTotal(componentType);
    //        return gff_total;
    //    }

    //    // Step 3 release hole size selection
    //    // 3.2
    //    public String releasePhase()
    //    {
    //        String release_phase = "";
    //        if ((p_s >= p_atm) && (p_s <= psat_s))
    //            release_phase = "Vapor";
    //        else if ((psat_s > p_atm) && (psat_s <= p_s))
    //            release_phase = "Two-phase";
    //        else
    //            release_phase = "Liquid";
    //        return release_phase;
    //    }

    //    // 3.3
    //    public float a_n(int n)
    //    {
    //        float a_n = 0;
    //        float dn = d_n(n);
    //        a_n = (float)(Math.PI * Math.Pow(dn, 2) / 4);
    //        return a_n;
    //    }
    //    // 3.4
    //    public float W_n(int n)
    //    {
    //        float W_n = 0;
    //        float an = a_n(n);
    //        String releasephase = releasePhase();
    //        float k = 0;
    //        float k_vn = 0;
    //        float mw = float.Parse(rbi.getMw(fluid));
    //        float p_trans = 0;
    //        if (releasephase.Equals("Liquid") || releasephase.Equals("Two-phase"))
    //        {
    //            k = (float)(0.9935 + 2.878 / Math.Pow(r_e, 0.5) + 342.75 / Math.Pow(r_e, 1.5));
    //            k_vn = (float)Math.Pow(k, -1);
    //            W_n = 0.61 * k_vn * rbi.getPl(fluid) * an * Math.Sqrt(2 * 6.67 * Math.Exp(-11) * Math.Abs(p_s - p_atm) / rbi.getPl(fluid)) / (rbi.getC(1));
    //        }
    //        else
    //        {
    //            float R = 8.314f;
    //            k = (Cp() / (Cp() - R));
    //            p_trans = (float)(p_atm * Math.Pow((k + 1) / 2, k / (k - 1)));
    //            if (p_s > p_trans)
    //            {
    //                float x = (k * mw * 6.67 * Math.Exp(-11) / (R * t_s)) * Math.Pow(2 / (k + 1), (k + 1) / (k - 1));
    //                W_n = 0.9 * an * p_s * Math.Sqrt(x) / (rbi.getC(2));
    //            }
    //            else
    //            {
    //                float x = (mw * 6.67 * Math.Exp(-11) / (8.314 * t_s)) * ((2 * k) / (k - 1)) * Math.Pow(p_atm / p_s, 2 / k) * (1 - Math.Pow(p_atm / p_s, (k - 1) / k));
    //                W_n = 0.9 * an * p_s * Math.Sqrt(x) / (rbi.getC(2));
    //            }
    //        }

    //        return W_n;
    //    }

    //    // cp
    //    public float Cp()
    //    {
    //        int idealCp = rbi.getCp_ideal(fluid);
    //        float cp = 0;
    //        float A = rbi.getCp(fluid, "A");
    //        float B = rbi.getCp(fluid, "B");
    //        float C = rbi.getCp(fluid, "C");
    //        float D = rbi.getCp(fluid, "D");
    //        float E = rbi.getCp(fluid, "E");
    //        float CP_C2 = (C / t_s) / (Math.Sinh(C / t_s));
    //        float CP_E2 = (E / t_s) / (Math.Cosh(E / t_s));
    //        if (idealCp == 1)
    //        {
    //            cp = A + B * t_s + C * t_s * t_s + D * t_s * t_s * t_s;
    //        }
    //        else if (idealCp == 2)
    //        {
    //            cp = A + B * CP_C2 * CP_C2 + D * CP_E2 * CP_E2;
    //        }
    //        else if (idealCp == 3)
    //        {
    //            cp = A + B * t_s + C * Math.Pow(t_s, 2) + D * Math.Pow(t_s, 3) + E * Math.Pow(t_s, 4);
    //        }
    //        else
    //        {
    //            cp = 0;
    //        }
    //        return cp;
    //    }


    //    // Step 4 : fluid inventory available for release
    //    // 4.5
    //    public float W_max8()
    //    {
    //        float W_max8 = 0;
    //        float an = 32450;
    //        String releasephase = releasePhase();
    //        float k = 0;
    //        float k_vn = 0;
    //        float mw = float.Parse(rbi.getMw(fluid));
    //        float p_trans = 0;
    //        if (releasephase.Equals("Liquid") || releasephase.Equals("Two-phase"))
    //        {
    //            k = 0.9935 + 2.878 / Math.Pow(r_e, 0.5) + 342.75 / Math.Pow(r_e, 1.5);
    //            k_vn = Math.Pow(k, -1);
    //            W_max8 = 0.61 * k_vn * rbi.getPl(fluid) * an * Math.Sqrt(2 * 6.67 * Math.Exp(-11) * Math.Abs(p_s - p_atm) / rbi.getPl(fluid)) / (rbi.getC(1));
    //        }
    //        else
    //        {
    //            float R = 8.314;
    //            k = (Cp() / (Cp() - R));
    //            p_trans = p_atm * Math.Pow((k + 1) / 2, k / (k - 1));
    //            if (p_s > p_trans)
    //            {
    //                float x = (k * mw * 6.67 * Math.Exp(-11) / (R * t_s)) * Math.Pow(2 / (k + 1), (k + 1) / (k - 1));
    //                W_max8 = 0.9 * an * p_s * Math.Sqrt(x) / (rbi.getC(2));
    //            }
    //            else
    //            {
    //                float x = (mw * 6.67 * Math.Exp(-11) / (8.314 * t_s)) * ((2 * k) / (k - 1)) * Math.Pow(p_atm / p_s, 2 / k) * (1 - Math.Pow(p_atm / p_s, (k - 1) / k));
    //                W_max8 = 0.9 * an * p_s * Math.Sqrt(x) / (rbi.getC(2));
    //            }
    //        }

    //        return W_max8;
    //    }
    //    // 4.6
    //    public float mass_addn(int n)
    //    {
    //        float mass_addn = 0;
    //        float Wmax8 = W_max8();
    //        float Wn = W_n(n);
    //        mass_addn = 180 * Math.Min(Wmax8, Wn);
    //        return mass_addn;
    //    }
    //    // 4.7
    //    public float mass_availn(int n)
    //    {
    //        float mass_availn = 0;
    //        float massaddn = mass_addn(n);
    //        mass_availn = Math.Min(mass_comp + massaddn, mass_inv);
    //        return mass_availn;
    //    }

    //    // Step 5 : releaseType
    //    // 5.1 
    //    public float t_n(int n)
    //    {
    //        float t_n = 0;
    //        float Wn = W_n(n);
    //        t_n = (rbi.getC(3)) / Wn;
    //        return t_n;
    //    }
    //    // 5.2
    //    public String releaseType(int n)
    //    {
    //        String releaseType = "";
    //        float tn = t_n(n);
    //        float dn = d_n(n);
    //        float massavailn = mass_availn(n);
    //        if (dn <= 6.35)
    //            releaseType = "Continuous";
    //        else if ((tn <= 180) || (massavailn > 4536))
    //            releaseType = "Instantaneous";
    //        else
    //            releaseType = "Continuous";
    //        return releaseType;
    //    }

    //    // Step 6 : Impact of detection and isolation system on release magnitude
    //    // 6.4
    //    public float fact_di()
    //    {
    //        float fact_di = 0;
    //        if (detectionType == 1 && isolationType == 1)
    //        {
    //            fact_di = 0.25;
    //        }
    //        else if (detectionType == 1 && isolationType == 2)
    //        {
    //            fact_di = 0.2;
    //        }
    //        else if ((detectionType == 1 || detectionType == 2) && isolationType == 3)
    //        {
    //            fact_di = 0.1;
    //        }
    //        else if (detectionType == 2 && isolationType == 2)
    //        {
    //            fact_di = 0.15;
    //        }
    //        else fact_di = 0;
    //        return fact_di;
    //    }

    //    //6.5
    //    public float ld_n_max(int n)
    //    {
    //        float ld_max = 0;
    //        float dn = d_n(n);
    //        if (detectionType == 1 && isolationType == 1)
    //        {
    //            if (dn == 0.25)
    //                ld_max = 20;
    //            else if (dn == 1)
    //                ld_max = 10;
    //            else //if (dn == 4)
    //                ld_max = 5;
    //        }
    //        else if (detectionType == 1 && isolationType == 2)
    //        {
    //            if (dn == 0.25)
    //                ld_max = 30;
    //            else if (dn == 1)
    //                ld_max = 20;
    //            else// if (dn == 4)
    //                ld_max = 10;
    //        }
    //        else if (detectionType == 1 && isolationType == 3)
    //        {
    //            if (dn == 0.25)
    //                ld_max = 40;
    //            else if (dn == 1)
    //                ld_max = 30;
    //            else// if (dn == 4)
    //                ld_max = 20;
    //        }
    //        else if ((isolationType == 1 || isolationType == 2) && detectionType == 2)
    //        {
    //            if (dn == 0.25)
    //                ld_max = 40;
    //            else if (dn == 1)
    //                ld_max = 30;
    //            else// if (dn == 4)
    //                ld_max = 20;
    //        }
    //        else if (detectionType == 2 && isolationType == 3)
    //        {
    //            if (dn == 0.25)
    //                ld_max = 60;
    //            else if (dn == 1)
    //                ld_max = 30;
    //            else// if (dn == 4)
    //                ld_max = 20;
    //        }
    //        else if (detectionType == 3 && (isolationType == 1 || isolationType == 2 || isolationType == 3))
    //        {
    //            if (dn == 0.25)
    //                ld_max = 60;
    //            else if (dn == 1)
    //                ld_max = 40;
    //            else// if (dn == 4)
    //                ld_max = 20;
    //        }
    //        else
    //            ld_max = 0;
    //        return ld_max;
    //    }


    //    // Step 7 : Release rate and consequence for analysis
    //    // 7.1
    //    public float rate_n(int n)
    //    {
    //        float rate_n = 0;
    //        float Wn = W_n(n);
    //        float factdi = fact_di();

    //        rate_n = Wn * (1 - factdi);
    //        return rate_n;
    //    }
    //    // 7.2
    //    public float ld_n(int n)
    //    {
    //        float ld_n = 0;
    //        float raten = rate_n(n);
    //        float massavailn = mass_availn(n);
    //        float ldmax = ld_n_max(n);
    //        ld_n = Math.Min(massavailn / raten, 60 * ldmax);
    //        return ld_n;
    //    }
    //    // 7.3
    //    public float frac_ro()
    //    {
    //        float frac_ro = 0;
    //        if (frac_fsh < 0.5)
    //            frac_ro = 1 - 2 * frac_fsh;
    //        else
    //            frac_ro = 0;
    //        return frac_ro;
    //    }
    //    // 7.4
    //    public float W_n_pool(int n)
    //    {
    //        float W_n_pool = 0;
    //        float raten = rate_n(n);
    //        float fracro = frac_ro();
    //        W_n_pool = raten * fracro;
    //        return W_n_pool;
    //    }
    //    // 7.5
    //    public float W_n_jet(int n)
    //    {
    //        float W_n_jet = 0;
    //        float raten = rate_n(n);
    //        float fracro = frac_ro();
    //        W_n_jet = raten * (1 - fracro);
    //        return W_n_jet;
    //    }
    //    // 7.6
    //    public float frac_entl()
    //    {
    //        float frac_entl = 0;
    //        float fracro = frac_ro();
    //        frac_entl = (frac_l * frac_fsh) / (1 - fracro);
    //        return frac_entl;
    //    }
    //    // 7.7
    //    public float erate_n(int n)
    //    {
    //        float erate_n = 0;
    //        String releasephase = releasePhase();
    //        float r_pn = Math.Sqrt(2 / 3) * Math.Pow((8 * 9.81 * V_pn) / (Math.PI), 0.25) * Math.Pow(t_pn, 0.75);
    //        float ksurf = rbi.getksurf(surface);
    //        float xsurf = rbi.getxsurf(surface);
    //        float alphasurf = rbi.getalphasurf(surface);

    //        if (releasephase.Equals("Vapor"))
    //            erate_n = W_n_jet(n);
    //        else// if (releasephase.Equals("Liquid"))
    //        {
    //            if (poolfireType.Equals("Non-boiling"))
    //                erate_n = (rbi.getC(15)) * P_bg * float.Parse(rbi.getMw(fluid)) * Math.Pow(r_pn, 1.89) * Math.Pow(u_w, 0.78);
    //            else// if (poolfireType.Equals("Boiling"))
    //                erate_n = Math.Pow(Math.PI, 1.5) * xsurf * ksurf * (t_g - t_b) * Math.Pow(2 * 9.81 * V_pn, 0.5) * t_pn / ((rbi.getC(14)) * 90 * Math.Sqrt(Math.PI * alphasurf));
    //        }
    //        return erate_n;
    //    }

    //    // Step 8 flamable explosive Consequences area
    //    // 8.2 flammable release rate
    //    public float rate_n_flame(int n)
    //    {
    //        return rate_n(n) * mfrac_flame;
    //    }
    //    public float rate_ln_flame(int n)
    //    {
    //        return rate_n_flame(n) * (1 - frac_fsh);
    //    }
    //    public float rate_vn_flame(int n)
    //    {
    //        return rate_n_flame(n) * frac_fsh;
    //    }


    //    // 8.4 probability of ignition of the release
    //    public float poi_ln_amb(int n)
    //    {
    //        return 1.00982 - 0.70372 * Math.Log(float.Parse(rbi.getMw(fluid))) - 0.013045 * Math.Log((rbi.getC(4)) * rate_ln_flame(n)) + 0.18554 * Math.Log(Math.Pow(float.Parse(rbi.getMw(fluid)), 2)) - 0.0014619 * Math.Log(Math.Pow((rbi.getC(4)) * rate_ln_flame(n), 2)) - 0.022131 * Math.Log(float.Parse(rbi.getMw(fluid))) * Math.Log((rbi.getC(4)) * rate_ln_flame(n)) - 0.016572 * Math.Log(Math.Pow(float.Parse(rbi.getMw(fluid)), 3)) + 0.00011281 * Math.Log(Math.Pow((rbi.getC(4)) * rate_ln_flame(n), 3)) + 0.00050697 * Math.Log(float.Parse(rbi.getMw(fluid))) * Math.Log(Math.Pow((rbi.getC(4)) * rate_ln_flame(n), 2)) - 0.0035535 * Math.Log(Math.Pow(float.Parse(rbi.getMw(fluid)), 2)) * Math.Log((rbi.getC(4)) * rate_ln_flame(n));

    //    }
    //    public float poi_vn_amb(int n)
    //    {
    //        return (1.16928 - 0.39309 * Math.Log(float.Parse(rbi.getMw(fluid))) - 0.053213 * Math.Log((rbi.getC(4)) * rate_vn_flame(n)) + 0.033904 * Math.Log(Math.Pow(float.Parse(rbi.getMw(fluid)), 2)) - 0.0028936 * Math.Log(Math.Pow((rbi.getC(4)) * rate_ln_flame(n), 2)) - 0.0067701 * Math.Log(float.Parse(rbi.getMw(fluid))) * Math.Log((rbi.getC(4)) * rate_ln_flame(n)) / (1 - 0.00110843 * Math.Log(float.Parse(rbi.getMw(fluid))) - 0.094276 * Math.Log((rbi.getC(4)) * rate_vn_flame(n)) + 0.029813 * Math.Log(Math.Pow(float.Parse(rbi.getMw(fluid)), 2)) + 0.0031951 * Math.Log(Math.Pow((rbi.getC(4)) * rate_ln_flame(n), 2)) - 0.058105 * Math.Log(float.Parse(rbi.getMw(fluid))) * Math.Log((rbi.getC(4)) * rate_vn_flame(n))));

    //    }
    //    public float poi_l_ait()
    //    {
    //        return 1.0;
    //    }
    //    public float poi_v_ait()
    //    {
    //        float max;
    //        float temp = (170 - float.Parse(rbi.getMw(fluid))) / (170 - 2);
    //        if (temp > 0)
    //            max = 0.7 + temp;
    //        else
    //            max = 0.7;
    //        return max;
    //    }
    //    public float poi_ln(int n)
    //    {
    //        return poi_ln_amb(n) + (poi_l_ait() - poi_ln_amb(n)) * (t_s - (rbi.getC(16))) / ((rbi.getAIT(fluid)) - (rbi.getC(16)));
    //    }
    //    public float poi_vn(int n)
    //    {
    //        return poi_vn_amb(n) + (poi_v_ait() - poi_vn_amb(n)) * (t_s - (rbi.getC(16))) / ((rbi.getAIT(fluid)) - (rbi.getC(16)));
    //    }
    //    public float poi_2n(int n)
    //    {
    //        return poi_ln(n) * frac_fsh + poi_vn(n) * (1 - frac_fsh);
    //    }


    //    // 8.5 probability of a immidiate ignition given ignition given immediate release 
    //    public float poii(String fp, int n)
    //    {
    //        float poii = 0;
    //        String releasetype = releaseType(n);
    //        if (releasetype.Equals("Continuous"))
    //        {
    //            if (fp.Equals("Liquid"))
    //            {
    //                if (ambientCondition.Equals("Ambient Temperature"))
    //                    poii = rbi.getpoii_n_amb_ambientTemperature(releasetype, fp);
    //                else// if (ambientCondition.Equals("AIT"))
    //                    poii = rbi.getpoii_AIT(releasetype, fp);
    //            }
    //            else// if (fp.Equals("Vapor"))
    //            {
    //                if (ambientCondition.Equals("Ambient Temperature"))
    //                    poii = rbi.getpoii_n_amb_ambientTemperature(releasetype, fp);
    //                else// if (ambientCondition.Equals("AIT"))
    //                    poii = rbi.getpoii_AIT(releasetype, fp);
    //            }
    //        }
    //        else// if (releasetype.Equals("Instantaneous"))
    //        {
    //            if (fp.Equals("Liquid"))
    //            {
    //                if (ambientCondition.Equals("Ambient Temperature"))
    //                    poii = rbi.getpoii_n_amb_ambientTemperature(releasetype, fp);
    //                else// if (ambientCondition.Equals("AIT"))
    //                    poii = rbi.getpoii_AIT(releasetype, fp);
    //            }
    //            else// if (fp.Equals("Vapor"))
    //            {
    //                if (ambientCondition.Equals("Ambient Temperature"))
    //                    poii = rbi.getpoii_n_amb_ambientTemperature(releasetype, fp);
    //                else// if (ambientCondition.Equals("AIT"))
    //                    poii = rbi.getpoii_AIT(releasetype, fp);
    //            }
    //        }
    //        return poii;
    //    }
    //    public float poii_ln(int n)
    //    {
    //        float poii_ln_amb = poii("Liquid", n);
    //        return poii_ln_amb + (1.0 - poii_ln_amb) * (t_s - (rbi.getC(16))) / ((rbi.getAIT(fluid)) - (rbi.getC(16)));
    //    }
    //    public float poii_vn(int n)
    //    {
    //        float poii_vn_amb = poii("Vapor", n);
    //        return poii_vn_amb + (1.0 - poii_vn_amb) * (t_s - (rbi.getC(16))) / ((rbi.getAIT(fluid)) - (rbi.getC(16)));
    //    }
    //    public float poii_2n(int n)
    //    {
    //        return frac_fsh * poii_ln(n) + (1 - frac_fsh) * poii_vn(n);
    //    }


    //    // 8.6 probability of a vce given delayed ignition
    //    public float pvce(String fp, int n)
    //    {
    //        String releasetype = releaseType(n);
    //        float pvce = rbi.getpvcedi(releasetype, fp);
    //        return pvce;
    //    }
    //    public float pvcedi_ln(int n)
    //    {
    //        return pvce("Liquid", n);
    //    }
    //    public float pvcedi_vn(int n)
    //    {
    //        return pvce("Vapor", n);
    //    }
    //    public float pvcedi_2n(int n)
    //    {
    //        return frac_fsh * pvcedi_ln(n) + (1 - frac_fsh) * pvcedi_vn(n);
    //    }


    //    // 8.7 probability of a flasf fire given delayed ignition 
    //    public float pffdi_ln(int n)
    //    {
    //        float pffdi_ln = 1 - pvcedi_ln(n);
    //        return pffdi_ln;
    //    }
    //    public float pffdi_vn(int n)
    //    {
    //        float pffdi_vn = 1 - pvcedi_vn(n);
    //        return pffdi_vn;
    //    }
    //    public float pffdi_2n(String fp, int n)
    //    {
    //        float pffdi_2n = frac_fsh * pffdi_ln(n) + (1 - frac_fsh) * pffdi_vn(n);
    //        return pffdi_2n;
    //    }


    //    // 8.8 probability of a fireball given immediate release 
    //    public float pfbii(String releasetype, String fluidPhase)
    //    {
    //        float pfbii = 0;
    //        if ((releasetype.Equals("Instantaneous") || fluidPhase.Equals("Vapor")) & releasetype.Equals("Instantaneous") || fluidPhase.Equals("Two-phase"))
    //            pfbii = 1;
    //        else
    //            pfbii = 0;
    //        return pfbii;
    //    }


    //    // 8.9 select appropriate event tree type
    //    public String eventTreeType(int n)
    //    {
    //        String eventTreeType = "";
    //        if ((n == 1) || (n == 2) || (n == 3))
    //            eventTreeType = "Leakage";
    //        else// if (n == 4)
    //            eventTreeType = "Rupture";
    //        return eventTreeType;
    //    }


    //    // 8.10 probability of all possible event outcomes
    //    public float p_poolfire(int n)
    //    {
    //        float p_poolFire = 0;
    //        String eTT = eventTreeType(n);
    //        String releasephase = releasePhase();
    //        String releasetype = releaseType(n);
    //        float poiln = poi_ln(n);
    //        float poiiln = poii_ln(n);
    //        float poi2n = poi_2n(n);
    //        float poii2n = poii_2n(n);
    //        float pvcedivn = pvcedi_vn(n);
    //        float pfbii2r = pfbii(releasetype, fluidPhase);
    //        if (eTT.Equals("Leakage"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                p_poolFire = 0;
    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_poolFire = poiln * poiiln;
    //            }
    //            else// if (releasephase.Equals("Two-phase"))
    //            {
    //                if (releasetype.Equals("Instantaneous"))
    //                    p_poolFire = poi2n * poii2n;
    //                else
    //                    p_poolFire = 0;

    //            }
    //        }
    //        else// if (eTT.Equals("Rupture"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                p_poolFire = 0;
    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_poolFire = poiln * poiiln;
    //            }
    //            else// if (releasephase.Equals("Two-phase"))
    //            {
    //                p_poolFire = poi2n * poii2n * (1 - pfbii2r);
    //            }
    //        }

    //        return p_poolFire;
    //    }
    //    public float p_jetfire(int n)
    //    {
    //        float p_jetfire = 0;
    //        String eTT = eventTreeType(n);
    //        String releasephase = releasePhase();
    //        String releasetype = releaseType(n);
    //        float poivn = poi_vn(n);
    //        float poiivn = poii_vn(n);
    //        float poi2n = poi_2n(n);
    //        float poii2n = poii_2n(n);

    //        if (eTT.Equals("Leakage"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                if (releasetype.Equals("Continuous"))
    //                    p_jetfire = poivn * poiivn;
    //                else
    //                    p_jetfire = 0;

    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_jetfire = 0;
    //            }
    //            else// if (releasephase.Equals("Two-phase"))
    //            {
    //                if (releasetype.Equals("Continuous"))
    //                    p_jetfire = poi2n * poii2n;
    //                else
    //                    p_jetfire = 0;
    //            }
    //        }
    //        else// if (eTT.Equals("Rupture"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                p_jetfire = 0;
    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_jetfire = 0;
    //            }
    //            else// if (releasephase.Equals("Two-phase"))
    //            {
    //                p_jetfire = 0;
    //            }
    //        }

    //        return p_jetfire;
    //    }
    //    public float p_fireball(int n)
    //    {
    //        float p_fireball = 0;
    //        String eTT = eventTreeType(n);
    //        String releasephase = releasePhase();
    //        String releasetype = releaseType(n);
    //        float poivn = poi_vn(n);
    //        float poiivn = poii_vn(n);
    //        float poi2n = poi_2n(n);
    //        float poii2n = poii_2n(n);
    //        float pfbii2r = pfbii(releasetype, fluidPhase);
    //        if (eTT.Equals("Leakage"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                if (releasetype.Equals("Instantaneous"))
    //                    p_fireball = poivn * poiivn;
    //                else
    //                    p_fireball = 0;
    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_fireball = 0;
    //            }
    //            else// if (releasephase.Equals("Two-phase"))
    //            {
    //                if (releasetype.Equals("Instantaneous"))
    //                    p_fireball = poi2n * poii2n;
    //                else
    //                    p_fireball = 0;
    //            }
    //        }
    //        else //if (eTT.Equals("Rupture"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                p_fireball = poivn * poiivn;
    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_fireball = 0;
    //            }
    //            else// if (releasephase.Equals("Two-phase"))
    //            {
    //                p_fireball = poi2n * poii2n * pfbii2r;
    //            }
    //        }

    //        return p_fireball;
    //    }
    //    public float p_vce(int n)
    //    {
    //        float p_vce = 0;
    //        String eTT = eventTreeType(n);
    //        String releasephase = releasePhase();
    //        String releasetype = releaseType(n);
    //        float poivn = poi_vn(n);
    //        float poiivn = poii_vn(n);
    //        float poiln = poi_ln(n);
    //        float poiiln = poii_ln(n);
    //        float poi2n = poi_2n(n);
    //        float poii2n = poii_2n(n);
    //        float pfbii2r = pfbii(releasetype, fluidPhase);
    //        float pvcedivn = pvcedi_vn(n);
    //        float pvcediln = pvcedi_ln(n);
    //        float pvcedi2n = pvcedi_2n(n);
    //        if (eTT.Equals("Leakage"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                p_vce = poivn * (1 - poiivn) * pvcedivn;
    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_vce = poiln * (1 - poiiln) * pvcediln;
    //            }
    //            else //if (releasephase.Equals("Two-phase"))
    //            {
    //                p_vce = poi2n * (1 - poii2n) * pvcedi2n;
    //            }
    //        }
    //        else// if (eTT.Equals("Rupture"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                p_vce = poivn * (1 - poiivn) * pvcedivn;
    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_vce = poiln * (1 - poiiln) * pvcediln;
    //            }
    //            else// if (releasephase.Equals("Two-phase"))
    //            {
    //                p_vce = poi2n * (1 - poii2n) * pvcedi2n;
    //            }
    //        }

    //        return p_vce;
    //    }
    //    public float p_flashfire(int n)
    //    {
    //        float p_flashfire = 0;
    //        String eTT = eventTreeType(n);
    //        String releasephase = releasePhase();
    //        String releasetype = releaseType(n);
    //        float poivn = poi_vn(n);
    //        float poiivn = poii_vn(n);
    //        float poiln = poi_ln(n);
    //        float poiiln = poii_ln(n);
    //        float poi2n = poi_2n(n);
    //        float poii2n = poii_2n(n);
    //        float pfbii2r = pfbii(releasetype, fluidPhase);
    //        float pvcedivn = pvcedi_vn(n);
    //        float pvcediln = pvcedi_ln(n);
    //        float pvcedi2n = pvcedi_2n(n);
    //        if (eTT.Equals("Leakage"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                p_flashfire = poivn * (1 - poiivn) * (1 - pvcedivn);
    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_flashfire = poiln * (1 - poiiln) * (1 - pvcediln);
    //            }
    //            else// if (releasephase.Equals("Two-phase"))
    //            {
    //                p_flashfire = poi2n * (1 - poii2n) * (1 - pvcedi2n);
    //            }
    //        }
    //        else// if (eTT.Equals("Rupture"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                p_flashfire = poivn * (1 - poiivn) * (1 - pvcedivn);
    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_flashfire = poiln * (1 - poiiln) * (1 - pvcediln);
    //            }
    //            else// if (releasephase.Equals("Two-phase"))
    //            {
    //                p_flashfire = poi2n * (1 - poii2n) * (1 - pvcedi2n);
    //            }
    //        }

    //        return p_flashfire;
    //    }
    //    public float p_safedispersion(int n)
    //    {
    //        float p_safedispersion = 0;
    //        String eTT = eventTreeType(n);
    //        String releasephase = releasePhase();
    //        float poivn = poi_vn(n);
    //        float poiln = poi_vn(n);
    //        float poi2n = poi_2n(n);

    //        if (eTT.Equals("Leakage"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                p_safedispersion = 1 - poivn;
    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_safedispersion = 1 - poiln;
    //            }
    //            else// if (releasephase.Equals("Two-phase"))
    //            {
    //                p_safedispersion = 1 - poi2n;
    //            }
    //        }
    //        else// if (eTT.Equals("Rupture"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                p_safedispersion = poivn;
    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_safedispersion = 1 - poiln;
    //            }
    //            else// if (releasephase.Equals("Two-phase"))
    //            {
    //                p_safedispersion = 0;
    //            }
    //        }

    //        return p_safedispersion;
    //    }
    //    public float p_pe(int n)
    //    {
    //        float p_pe = 0;
    //        String eTT = eventTreeType(n);
    //        String releasephase = releasePhase();
    //        float poivn = poi_vn(n);

    //        if (eTT.Equals("Leakage"))
    //        {
    //            p_pe = 0;
    //        }
    //        else// if (eTT.Equals("Rupture"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                p_pe = 1 - poivn;
    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_pe = 0;
    //            }
    //            else// if (releasephase.Equals("Two-phase"))
    //            {
    //                p_pe = 0;
    //            }
    //        }
    //        return p_pe;
    //    }
    //    public float p_bleve(int n)
    //    {
    //        float p_bleve = 0;
    //        String eTT = eventTreeType(n);
    //        String releasephase = releasePhase();
    //        float poi2n = poi_2n(n);

    //        if (eTT.Equals("Leakage"))
    //        {
    //            p_bleve = 0;
    //        }
    //        else// if (eTT.Equals("Rupture"))
    //        {
    //            if (releasephase.Equals("Vapor"))
    //            {
    //                p_bleve = 0;
    //            }
    //            else if (releasephase.Equals("Liquid"))
    //            {
    //                p_bleve = 0;
    //            }
    //            else// if (releasephase.Equals("Two-phase"))
    //            {
    //                p_bleve = 1 - poi2n;
    //            }
    //        }
    //        return p_bleve;
    //    }

    //    // 8.11 component damage and personal injury of fire pool
    //    public float m_b()
    //    {
    //        float m_b = 0;
    //        if (poolfireType.Equals("Non-boiling"))
    //            m_b = ((rbi.getC(17)) * 45 * Math.Pow(10, 6)) / (c_pl * (t_b - t_atm) + 20 * Math.Pow(10, 5));
    //        else// if (poolfireType.Equals("Boiling"))
    //            m_b = ((rbi.getC(17)) * 45 * Math.Pow(10, 6)) / (20 * Math.Pow(10, 5));
    //        return m_b;
    //    }
    //    public float aburn_pfn(int n)
    //    {
    //        float aburn_pfn = 0;
    //        aburn_pfn = W_n_pool(n) / m_b();
    //        return aburn_pfn;
    //    }
    //    public float a_max_pfn(int n)
    //    {
    //        float a_max_pfn = 0;
    //        a_max_pfn = mass_availn(n) / ((rbi.getC(18)) * frac_ro() * p_l);
    //        return a_max_pfn;
    //    }
    //    public float a_pfn(int n)
    //    {
    //        float a_pfn = 0;
    //        float amaxpfn = a_max_pfn(n);
    //        float aburnpfn = aburn_pfn(n);
    //        a_pfn = Math.Min(Math.Min(amaxpfn, aburnpfn), (rbi.getC(7)) * 929.1);
    //        return a_pfn;
    //    }
    //    public float r_pfn(int n)
    //    {
    //        float r_pfn = 0;
    //        float apfn = a_pfn(n);

    //        r_pfn = Math.Sqrt(apfn / Math.PI);
    //        return r_pfn;
    //    }
    //    public float u_sn(int n)
    //    {
    //        float u_sn = 0;
    //        float mb = m_b();
    //        float rpfn = r_pfn(n);
    //        u_sn = Math.Max(1.0, u_w * Math.Pow(p_v / (2 * 9.81 * mb * rpfn), 0.333));
    //        return u_sn;
    //    }
    //    public float l_pfn(int n)
    //    {
    //        float l_pfn = 0;
    //        float mb = m_b();
    //        float rpfn = r_pfn(n);
    //        float usn = u_sn(n);
    //        l_pfn = 110 * rpfn * Math.Pow(mb / (p_atm * Math.Sqrt(2 * 9.81 * rpfn)), 0.67) * Math.Pow(usn, -0.21);
    //        return l_pfn;
    //    }
    //    public float O_pfn(int n)
    //    {
    //        float O_pfn = 0;
    //        float usn = u_sn(n);
    //        O_pfn = Math.Acos(1 / usn);
    //        return O_pfn;
    //    }
    //    public float qrad_n_pool(int n)
    //    {
    //        float qrad_n_pool = 0;
    //        float lpfn = l_pfn(n);
    //        float rpfn = r_pfn(n);
    //        float mb = m_b();
    //        qrad_n_pool = ((rbi.getC(14)) * 0.35 * mb * 45 * Math.Pow(10, 6) * Math.PI * Math.Pow(rpfn, 2)) / (2 * Math.PI * rpfn * lpfn + Math.PI * Math.Pow(rpfn, 2));
    //        return qrad_n_pool;
    //    }
    //    public float t_aimn(float xs_n)
    //    {
    //        float t_aimn = 0;
    //        float pw = (rbi.getC(20)) * RH * Math.Exp(14.4114 - (rbi.getC(19)) / t_atm);

    //        t_aimn = 0.819 * Math.Pow(pw * xs_n, -0.09);
    //        return t_aimn;
    //    }
    //    public float f_cyln(float xs_n_pool, int n)
    //    {
    //        float f_cyln = 0;
    //        float lpfn = l_pfn(n);
    //        float rpfn = r_pfn(n);
    //        float opfn = O_pfn(n);
    //        float X = lpfn / rpfn;
    //        float Y = xs_n_pool / rpfn;
    //        float A = Math.Pow(X, 2) + Math.Pow(Y + 1, 2) - 2 * X * (Y + 1) * Math.Sin(opfn);
    //        float B = Math.Pow(X, 2) + Math.Pow(Y - 1, 2) - 2 * X * (Y - 1) * Math.Sin(opfn);
    //        float C = 1 + (Math.Pow(Y, 2) - 1) * Math.Pow(Math.Cos(opfn), 2);
    //        float fvn = (X * Math.Cos(opfn)) / (Y - X * Math.Sin(opfn)) * (Math.Pow(X, 2) + Math.Pow(Y + 1, 2) - 2 * Y * (1 + Math.Sin(opfn))) / (Math.PI * Math.Sqrt(A * B)) * Math.Atan((A * (Y - 1)) / (B * (Y + 1))) + (Math.Cos(opfn)) / (Math.PI * Math.Sqrt(C)) * (Math.Atan((X * Y - (Math.Pow(Y, 2) * Math.Sin(opfn))) / (Math.Sqrt(Math.Pow(Y, 2) - 1) * Math.Sqrt(C))) + Math.Atan((Math.Sin(opfn) * Math.Sqrt(Math.Pow(Y, 2) - 1)) / (Math.Sqrt(C)))) - (X * Math.Cos(opfn)) / (Math.PI * (Y - X * Math.Sin(opfn))) * Math.Atan(Math.Sqrt((Y - 1) / (Y + 1)));
    //        float fhn = 1 / Math.PI * Math.Atan(Math.Sqrt((Y + 1) / (Y - 1))) - (Math.Pow(X, 2) + Math.Pow(Y + 1, 2) - 2 * (Y + 1 + X * Y * Math.Sin(opfn))) / (Math.PI * Math.Sqrt(A * B)) * Math.Atan(Math.Sqrt((A * (Y - 1)) / (B * (Y + 1)))) + (Math.Sin(opfn)) / (Math.PI * Math.Sqrt(C)) * (Math.Atan((X * Y - (Math.Pow(Y, 2) - 1) * Math.Sin(opfn)) / (Math.Sqrt(Math.Pow(Y, 2) - 1) * Math.Sqrt(C)) + Math.Atan((Math.Sin(opfn) * Math.Sqrt(Math.Pow(Y, 2) - 1)) / (Math.Sqrt(C)))));

    //        f_cyln = Math.Sqrt(Math.Pow(fvn, 2) + Math.Pow(fhn, 2));
    //        return f_cyln;
    //    }
    //    public float ith_n_pool(float xs_n_pool, int n)
    //    {
    //        float ith_n_pool = 0;
    //        float tatmn = t_aimn(xs_n_pool);
    //        float qradnpool = qrad_n_pool(n);
    //        float fcyln = f_cyln(xs_n_pool, n);

    //        ith_n_pool = (rbi.getC(19)) * qradnpool * tatmn * fcyln;

    //        return ith_n_pool;
    //    }
    //    public float ca_pool(float xs_n_pool)
    //    {
    //        float ca_pool = 0;
    //        ca_pool = Math.PI * Math.Pow(xs_n_pool, 2);
    //        return ca_pool;
    //    }

    //    // 8.12 component damage and personal injury of jet fire
    //    public float qrad_n_jet(int n)
    //    {
    //        float qrad_n_jet = 0;
    //        float mb = m_b();
    //        qrad_n_jet = (rbi.getC(14)) * 0.35 * W_n_jet(n) * 55.50 * Math.Pow(10, 6);
    //        return qrad_n_jet;
    //    }
    //    public float fp_n(float xs_n_jet)
    //    {
    //        float fp_n = 0;
    //        fp_n = 1 / (4 * Math.PI * Math.Pow(xs_n_jet, 2));
    //        return fp_n;
    //    }
    //    public float ith_n_jet(float xs_n_jet, int n)
    //    {

    //        float ith_n_jet = 0;
    //        float tatmn = t_aimn(xs_n_jet);
    //        float qradnjet = qrad_n_jet(n);
    //        float fpn = fp_n(xs_n_jet);


    //        ith_n_jet = tatmn * qradnjet * fpn;

    //        return ith_n_jet;
    //    }
    //    public float ca_jet(float xs_n_jet)
    //    {
    //        float ca_jet = 0;
    //        ca_jet = Math.PI * Math.Pow(xs_n_jet, 2);
    //        return ca_jet;
    //    }

    //    // 8.13 component damage and personal injury of fireball
    //    public float mass_fb(int n)
    //    {
    //        float mass_fb = 0;
    //        mass_fb = mfrac_flame * mass_availn(n);
    //        return mass_fb;
    //    }
    //    public float d_max_fb(int n)
    //    {
    //        float d_max_fb = 0;
    //        float massfb = mass_fb(n);
    //        d_max_fb = (rbi.getC(22)) * Math.Pow(massfb, 0.333);
    //        return d_max_fb;
    //    }
    //    public float h_fb(int n)
    //    {
    //        float h_fb = 0;
    //        float dmaxfb = d_max_fb(n);
    //        h_fb = 0.75 * dmaxfb;
    //        return h_fb;
    //    }
    //    public float t_fb(int n)
    //    {
    //        float t_fb = 0;
    //        float massfb = mass_fb(n);
    //        if (massfb < 29.937)
    //            t_fb = (rbi.getC(23)) * Math.Pow(massfb, 0.333);
    //        else
    //            t_fb = (rbi.getC(24)) * Math.Pow(massfb, 0.167);
    //        return t_fb;
    //    }
    //    public float qrad_fball(int n)
    //    {
    //        float qrad_fball = 0;
    //        float massfb = mass_fb(n);
    //        float dmax = d_max_fb(n);
    //        float tfb = t_fb(n);
    //        float Bfb = (rbi.getC(25)) * Math.Pow(p_s, 0.32);
    //        qrad_fball = ((rbi.getC(14)) * Bfb * massfb * 45 * Math.Pow(10, 6)) / (Math.PI * Math.Pow(dmax, 2) * tfb);
    //        return qrad_fball;
    //    }
    //    public float fsph(float xs_n_fball, int n)
    //    {
    //        float fsph = 0;
    //        float dmax = d_max_fb(n);
    //        float cfb = Math.Sqrt(Math.Pow(dmax / 2, 2) + Math.Pow(xs_n_fball / 2, 2));
    //        fsph = Math.Pow(dmax, 2) / (4 * Math.Pow(cfb, 2));
    //        return fsph;
    //    }
    //    public float ith_n_fb(float xs_n_fb, int n)
    //    {
    //        float ith_n_fb = 0;
    //        float tatmn = t_aimn(xs_n_fb);
    //        float qradnjet = qrad_fball(n);
    //        float fsph1 = fsph(xs_n_fb, n);
    //        ith_n_fb = tatmn * qradnjet * fsph1;

    //        return ith_n_fb;
    //    }
    //    public float ca_fb(float xs_fball)
    //    {
    //        float ca_fb = 0;
    //        ca_fb = Math.PI * Math.Pow(xs_fball, 2);
    //        return ca_fb;
    //    }

    //    // 8.14 component damage and personal injury of vce
    //    public float W_tnt()
    //    {
    //        float W_tnt = 0;
    //        W_tnt = (0.09 * mass_vce * 90 * Math.Pow(10, 6)) / (4648);
    //        return W_tnt;
    //    }
    //    public float R_hsn(float xs_n_vce)
    //    {
    //        float R_hsn = 0;
    //        float wtnt = W_tnt();
    //        R_hsn = ((rbi.getC(27)) * xs_n_vce) / (Math.Pow(wtnt, 0.3333333333));
    //        return R_hsn;
    //    }
    //    public float P_son(float xs_n_vce)
    //    {
    //        float P_son = 0;
    //        float rhsn = R_hsn(xs_n_vce);
    //        P_son = Math.Abs(-0.059965896 + (1.1288697) / (Math.Log(rhsn)) - (7.9625216) / (Math.Pow(Math.Log(rhsn), 2)) + (25.106738) / (Math.Pow(Math.Log(rhsn), 3)) - (30.396707) / (Math.Pow(Math.Log(rhsn), 4)) + (19.399862) / (Math.Pow(Math.Log(rhsn), 5)) - (6.8853477) / (Math.Pow(Math.Log(rhsn), 6)) + (1.2825511) / (Math.Pow(Math.Log(rhsn), 7)) - (0.097705789) / (Math.Pow(Math.Log(rhsn), 8)));
    //        return P_son;
    //    }

    //    public float ca_vce(float xs_n_vce)
    //    {
    //        float ca_vce = 0;
    //        ca_vce = Math.PI * Math.Pow(xs_n_vce, 2);
    //        return ca_vce;
    //    }
    //    public float p_r(float xs_n_vce)
    //    {
    //        float p_r = 0;
    //        float pson = P_son(xs_n_vce);
    //        p_r = Math.Abs(-23.8 + 2.92 * Math.Log((rbi.getC(28)) * pson));
    //        return p_r;
    //    }

    //    // 8.15 component damage and personal injury of flash fire
    //    public float ca_cmdn_flash()
    //    {
    //        float ca_cmdn_flash = 0;
    //        ca_cmdn_flash = 0.25 * gradecloudArea;
    //        return ca_cmdn_flash;
    //    }

    //    // 8.16
    //    public float ca_cmdn_flame(int n)
    //    {
    //        float ca_cmdn_flame = 0;
    //        ca_cmdn_flame = p_poolfire(n) * ca_pool(xs_cmdn_pool) + p_jetfire(n) * ca_jet(xs_cmdn_jet) + p_fireball(n) * ca_fb(xs_cmdn_fb) + p_vce(n) * ca_vce(xs_cmdn_vce) + p_flashfire(n) * ca_cmdn_flash();
    //        return ca_cmdn_flame;
    //    }
    //    public float ca_injn_flame(int n)
    //    {
    //        float ca_injn_flame = 0;
    //        ca_injn_flame = p_poolfire(n) * ca_pool(xs_injn_pool) + p_jetfire(n) * ca_jet(xs_injn_jet) + p_fireball(n) * ca_fb(xs_injn_fb) + p_vce(n) * ca_vce(xs_injn_vce) + p_flashfire(n) * gradecloudArea;
    //        return ca_injn_flame;
    //    }

    //    //8.17
    //    public float ca_cmd_flame()
    //    {
    //        float ca_cmd_flame = 0;
    //        float t = 0;
    //        for (int i = 0; i < 5; i++)
    //        {
    //            t += gff_n(i) * ca_cmdn_flame(i);
    //        }
    //        ca_cmd_flame = t / gff_total();
    //        return ca_cmd_flame;
    //    }
    //    public float ca_inj_flame()
    //    {
    //        float ca_inj_flame = 0;
    //        float t = 0;
    //        for (int i = 0; i < 5; i++)
    //        {
    //            t += gff_n(i) * ca_injn_flame(i);
    //        }
    //        ca_inj_flame = t / gff_total();
    //        return ca_inj_flame;
    //    }



    //    // Step 9 toxic consequences area
    //    public float mass_n(int n)
    //    {
    //        float mass_n = 0;
    //        mass_n = Math.Min(rate_n(n) * ld_n(n), mass_availn(n));

    //        return mass_n;
    //    }

    //    // 9.2
    //    public float ld_n_tox(int n)
    //    {
    //        float ld_n_tox = 0;
    //        ld_n_tox = Math.Min(Math.Min(3600, (mass_n(n)) / (W_n(n))), 60 * ld_n_max(n));
    //        return ld_n_tox;
    //    }
    //    // 9.3
    //    public float tox_lim()
    //    {
    //        float tox_lim = 0;
    //        tox_lim = rbi.getToxicImpactCriteria(toxicComponent, criteria);
    //        return tox_lim;
    //    }
    //    // 9.4
    //    public float tox_lim_mod()
    //    {
    //        float tox_lim_mod = 0;
    //        float toxlim = tox_lim();
    //        tox_lim_mod = toxlim / molefrac_tox;
    //        return tox_lim_mod;
    //    }
    //    // 9.7
    //    public float ptox_n(int n)
    //    {
    //        float ptox_n = 0;
    //        ptox_n = p_safedispersion(n);
    //        return ptox_n;
    //    }
    //    public float ca_injn_tox(int n)
    //    {
    //        float ca_n_tox = 0;
    //        float pntox = ptox_n(n);
    //        ca_n_tox = pntox * ca_n_cloud;
    //        return ca_n_tox;
    //    }
    //    public float ca_tox()
    //    {
    //        float ca_tox = 0;
    //        float t = 0;
    //        for (int i = 0; i < 5; i++)
    //        {
    //            t += gff_n(i) * ca_injn_tox(i);
    //        }
    //        ca_tox = t / gff_total();
    //        return Math.Abs(ca_tox);
    //    }


    //    // Step 10 non flammable non toxic consequence area
    //    // 10.2
    //    public float ca_injn_cont(int n)
    //    {
    //        float ca_injn_cont = 0;
    //        float h = 0.31 - 0.0032 * Math.Pow((rbi.getC(11)) * (p_s - p_atm) - 40, 2);
    //        float g = 2696 - 21.9 * (rbi.getC(11)) * (p_s - p_atm) + 1.474 * Math.Pow((rbi.getC(11)) * (p_s - p_atm), 2);
    //        if (representFluid.Equals("Steam"))
    //            ca_injn_cont = (rbi.getC(9)) * rate_n(n);
    //        else// if (representFluid.Equals("Acids") || representFluid.Equals("Caustics"))
    //            ca_injn_cont = 0.2 * (rbi.getC(8)) * g * Math.Pow((rbi.getC(4)) * rate_n(n), h);

    //        return ca_injn_cont;
    //    }
    //    public float ca_injn_inst(int n)
    //    {
    //        float ca_injn_inst = 0;

    //        if (representFluid.Equals("Steam"))
    //            ca_injn_inst = (rbi.getC(10)) * Math.Pow(mass_n(n), 0.6384);
    //        else// if (representFluid.Equals("Acids") || representFluid.Equals("Caustics"))
    //            ca_injn_inst = 0;

    //        return ca_injn_inst;
    //    }
    //    public float fact_n_ic(int n)
    //    {
    //        if (representFluid.Equals("Steam"))
    //            return Math.Min(rate_n(n) / (rbi.getC(5)), 1);
    //        else// if (representFluid.Equals("Acids") || representFluid.Equals("Caustics"))
    //            return 0;
    //    }
    //    public float ca_injn_leak(int n)
    //    {
    //        float ca_injn_leak = 0;
    //        float cainjninst = ca_injn_inst(n);
    //        float factnic = fact_n_ic(n);
    //        float cainjncont = ca_injn_cont(n);
    //        ca_injn_leak = cainjninst * factnic + cainjncont * (1 - factnic);

    //        return ca_injn_leak;
    //    }

    //    // 10.2
    //    public float ca_cmd_pexp()
    //    {
    //        float ca_cmd_pexp = 0;
    //        String eventtreeType = eventTreeType(4);
    //        if (eventtreeType.Equals("Rupture"))
    //        {
    //            ca_cmd_pexp = Math.PI * Math.Pow(xs_cmdn_pexp, 2);
    //        }
    //        return ca_cmd_pexp;
    //    }
    //    public float ca_inj_pexp()
    //    {
    //        float ca_inj_pexp = 0;
    //        String eventtreeType = eventTreeType(4);
    //        if (eventtreeType.Equals("Rupture"))
    //        {
    //            ca_inj_pexp = Math.PI * Math.Pow(xs_injn_pexp, 2);
    //        }
    //        return ca_inj_pexp;
    //    }

    //    // 10.3
    //    public float W_tnt_bleve()
    //    {
    //        float Wtntbleve = 0;
    //        String releasephase = releasePhase();
    //        Wtntbleve = (rbi.getC(30)) * n_v * 8.314 * t_s * Math.Log(p_s / p_atm);
    //        float Wtnt = W_tnt();
    //        if (releasephase.Equals("Two-phase"))
    //            return Wtntbleve + Wtnt;
    //        else
    //            return Wtntbleve;
    //        //return Wtntbleve;
    //    }
    //    public float ca_cmd_bleve()
    //    {
    //        float ca_cmd_bleve = 0;
    //        ca_cmd_bleve = Math.PI * Math.Pow(xs_cmdn_bleve, 2);
    //        return ca_cmd_bleve;
    //    }

    //    public float ca_inj_bleve()
    //    {
    //        float ca_inj_bleve = 0;
    //        String eventtreeType = eventTreeType(4);
    //        if (eventtreeType.Equals("Rupture"))
    //        {
    //            ca_inj_bleve = Math.PI * Math.Pow(xs_injn_bleve, 2);
    //        }
    //        return ca_inj_bleve;
    //    }
    //    // 10.4
    //    public float ca_cmdn_nfnt()
    //    {
    //        float ca_cmdn_nfnt = 0;
    //        float cacmdpexp = ca_cmd_pexp();
    //        float cainjbleve = ca_cmd_bleve();

    //        ca_cmdn_nfnt = Math.Max(cacmdpexp, cainjbleve);
    //        return ca_cmdn_nfnt;
    //    }
    //    public float ca_injn_nfnt(int n)
    //    {
    //        float ca_injn_nfnt = 0;
    //        float cacmdpexp = ca_inj_pexp();
    //        float cainjbleve = ca_inj_bleve();
    //        float cainjleak = ca_injn_leak(n);

    //        ca_injn_nfnt = Math.Max(cacmdpexp, cainjbleve) + cainjleak;
    //        return ca_injn_nfnt;
    //    }
    //    // 10.5
    //    public float ca_cmd_nfnt()
    //    {
    //        float ca_cmd_nfnt = 0;
    //        float t = 0;
    //        for (int i = 0; i < 5; i++)
    //        {
    //            t += gff_n(i) * ca_cmdn_nfnt();
    //        }
    //        ca_cmd_nfnt = t / gff_total();
    //        return ca_cmd_nfnt;
    //    }
    //    public float ca_inj_nfnt()
    //    {
    //        float ca_inj_nfnt = 0;
    //        float t = 0;
    //        for (int i = 0; i < 5; i++)
    //        {
    //            t += gff_n(i) * ca_injn_nfnt(i);
    //        }
    //        ca_inj_nfnt = t / gff_total();
    //        return ca_inj_nfnt;
    //    }

    //    // Step 11 component and injury consequences
    //    // 11.1
    //    public float ca_cmd()
    //    {
    //        float ca_cmd = 0;
    //        float cacmdflame = ca_cmd_flame();
    //        float cacmdntft = ca_cmd_nfnt();
    //        ca_cmd = Math.Max(cacmdflame, cacmdntft);
    //        return ca_cmd;
    //    }

    //    // 11.2
    //    public float ca_inj()
    //    {
    //        float ca_inj = 0;
    //        float cainjflame = ca_inj_flame();
    //        float cainjnfnt = ca_inj_nfnt();
    //        float cainjtox = ca_tox();
    //        ca_inj = Math.Max(Math.Max(cainjflame, cainjtox), cainjnfnt);
    //        return ca_inj;
    //    }

    //    // Step 12: financial consequence
    //    // 12.1
    //    public float fc_cmd()
    //    {
    //        float fc_cmd = 0;
    //        float matcost = rbi.getMatcost(material);
    //        float t = 0;
    //        for (int i = 1; i < 5; i++)
    //        {
    //            t += gff_n(i) * rbi.getHoleSize(componentType, i) * matcost;
    //        }
    //        fc_cmd = t / gff_total();
    //        return fc_cmd;
    //    }

    //    //12.2
    //    public float fc_affa()
    //    {
    //        float fc_affa = 0;
    //        float cacmd = ca_cmd();

    //        fc_affa = cacmd * equipcost;
    //        return fc_affa;
    //    }

    //    // 12.3
    //    public float outage_cmd()
    //    {
    //        float t = 0;
    //        for (int i = 0; i < 5; i++)
    //        {
    //            t += gff_n(i) * rbi.getOutage(componentType, i) * outage_mult;
    //        }
    //        return t / gff_total();
    //    }
    //    public float outage_affa()
    //    {
    //        float outage_affa = 0;
    //        float fcaffa = fc_affa();

    //        outage_affa = Math.Pow(10, 1.242 + 0.585 * Math.Log10(fcaffa * Math.Pow(10, -6)));
    //        return outage_affa;
    //    }
    //    public float fc_prod()
    //    {
    //        float fc_prod = 0;
    //        float outagecmd = outage_cmd();
    //        float outageaffa = outage_affa();

    //        fc_prod = (outagecmd + outageaffa) * prodcost;
    //        return fc_prod;
    //    }

    //    // 12.4
    //    public float fc_inj()
    //    {
    //        float fc_inj = 0;
    //        float cainj = ca_inj();
    //        float outageaffa = outage_affa();

    //        fc_inj = cainj * popdens * injcost;
    //        return fc_inj;
    //    }

    //    // 12.5
    //    public float vol_n_env(int n)
    //    {
    //        float vol_n_env = 0;
    //        float massn = mass_n(n);
    //        float frac_evap = rbi.getfracEvap(fluid);

    //        vol_n_env = (rbi.getC(13)) * massn * (1 - frac_evap) / p_l;
    //        return vol_n_env;
    //    }
    //    public float fc_environ()
    //    {
    //        float fc_environ = 0;

    //        float t = 0;
    //        for (int i = 0; i < 5; i++)
    //        {
    //            float volnenv = vol_n_env(i);
    //            t += gff_n(i) * volnenv * envcost;
    //        }

    //        fc_environ = t / gff_total();
    //        return fc_environ;
    //    }


    //    // 12.6
    //    public float fc()
    //    {
    //        float fc = 0;
    //        float fccmd = fc_cmd();
    //        float fcaffa = fc_affa();
    //        float fcprod = fc_prod();
    //        float fcinj = fc_inj();
    //        float fcenviron = fc_environ();
    //        fc = fccmd + fcaffa + fcprod + fcinj + fcenviron;
    //        return fc;
    //    }

    //    // consequence analysis
    //    public float ca()
    //    {
    //        return ca_cmd() + ca_inj();
    //    }

    //    // catalog
    //    public int convertCatalog()
    //    {
    //        float c_a = ca();
    //        if (c_a <= 100)
    //            return 1;
    //        else if (c_a > 100 && c_a <= 1000)
    //            return 2;
    //        else if (c_a > 1000 && c_a <= 3000)
    //            return 3;
    //        else if (c_a > 3000 && c_a <= 10000)
    //            return 4;
    //        else return 5;
    //    }
    //}
}
