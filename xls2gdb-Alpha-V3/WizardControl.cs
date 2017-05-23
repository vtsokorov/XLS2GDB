using System;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace xls2gdb
{
    public partial class WizardControl : TabControl
    {
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x1328 && !DesignMode)
                m.Result = (IntPtr)1;
            else
                base.WndProc(ref m);
        }
    }
}
