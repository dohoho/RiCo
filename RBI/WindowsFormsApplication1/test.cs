﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using RBI.DAL.MSSQL_CAL;
using System.Diagnostics;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Design;
using DevExpress.XtraTreeList.Menu;
using DevExpress.XtraTreeList.Nodes;
using System.Net;
using RBI.BUS.BUSExcel;
using RBI.Object.ObjectMSSQL;
using DevExpress.XtraCharts;
using RBI.BUS.BUSMSSQL_CAL;

namespace RBI
{
    
    public partial class test : Form
    {
        MSSQL_DM_CAL d = new MSSQL_DM_CAL();
        public test()
        {
            InitializeComponent();
            //thinning

        }

    }
}
