﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RBI.Object.ObjectMSSQL
{
    public class FACILITY_RISK_TARGET
    {
        public int FacilityID { set; get; }
        public float RiskTarget_A { set; get; }
        public float RiskTarget_B { set; get; }
        public float RiskTarget_C { set; get; }
        public float RiskTarget_D { set; get; }
        public float RiskTarget_E { set; get; }
        public float RiskTarget_CA { set; get; }
        public float RiskTarget_FC { set; get; }
    }
}
