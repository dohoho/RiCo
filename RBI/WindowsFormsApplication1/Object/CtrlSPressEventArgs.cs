using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RBI.Object
{
    public class CtrlSPressEventArgs : EventArgs
    {
        public readonly int ctrlSpress;
        public CtrlSPressEventArgs(int p)
        {
            this.ctrlSpress = p;
        }

    }
}
