﻿using RBI.DAL.MSSQL;
using RBI.Object;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RBI.Object.ObjectMSSQL;
namespace RBI.BUS.BUSMSSQL
{
    class RW_EXTRA_FIELDS_BUS
    {
        RW_EXTRA_FIELDS_ConnetUtils DAL = new RW_EXTRA_FIELDS_ConnetUtils();
        public void add(RW_EXTRA_FIELDS obj)
        {
            DAL.add(obj.ID, obj.ObjectField001, obj.ObjectField002, obj.ObjectField003, obj.ObjectField004, obj.ObjectField005, obj.ObjectField006, obj.ObjectField007, obj.ObjectField008, obj.ObjectField009, obj.ObjectField010, obj.ObjectField011, obj.ObjectField012, obj.ObjectField013, obj.ObjectField014,
            obj.ObjectField015, obj.ObjectField016, obj.ObjectField017, obj.ObjectField018, obj.ObjectField019, obj.ObjectField020, obj.ObjectField021, obj.ObjectField022, obj.ObjectField023, obj.ObjectField024, obj.ObjectField025, obj.ObjectField026, obj.ObjectField027, obj.ObjectField028, obj.ObjectField029,
            obj.ObjectField030, obj.ObjectField031, obj.ObjectField032, obj.ObjectField033, obj.ObjectField034, obj.ObjectField035, obj.ObjectField036, obj.ObjectField037, obj.ObjectField038, obj.ObjectField039, obj.ObjectField040, obj.ObjectField041, obj.ObjectField042, obj.ObjectField043, obj.ObjectField044,
            obj.ObjectField045, obj.ObjectField046, obj.ObjectField047, obj.ObjectField048, obj.ObjectField049, obj.ObjectField050);
            
        }
        public void edit(RW_EXTRA_FIELDS obj)
        {
            DAL.edit(obj.ID, obj.ObjectField001, obj.ObjectField002, obj.ObjectField003, obj.ObjectField004, obj.ObjectField005, obj.ObjectField006, obj.ObjectField007, obj.ObjectField008, obj.ObjectField009, obj.ObjectField010, obj.ObjectField011, obj.ObjectField012, obj.ObjectField013, obj.ObjectField014,
            obj.ObjectField015, obj.ObjectField016, obj.ObjectField017, obj.ObjectField018, obj.ObjectField019, obj.ObjectField020, obj.ObjectField021, obj.ObjectField022, obj.ObjectField023, obj.ObjectField024, obj.ObjectField025, obj.ObjectField026, obj.ObjectField027, obj.ObjectField028, obj.ObjectField029,
            obj.ObjectField030, obj.ObjectField031, obj.ObjectField032, obj.ObjectField033, obj.ObjectField034, obj.ObjectField035, obj.ObjectField036, obj.ObjectField037, obj.ObjectField038, obj.ObjectField039, obj.ObjectField040, obj.ObjectField041, obj.ObjectField042, obj.ObjectField043, obj.ObjectField044,
            obj.ObjectField045, obj.ObjectField046, obj.ObjectField047, obj.ObjectField048, obj.ObjectField049, obj.ObjectField050);
            
        }
        public void delete(RW_EXTRA_FIELDS obj)
        {
            DAL.delete(obj.ID);
        }
    }
}
