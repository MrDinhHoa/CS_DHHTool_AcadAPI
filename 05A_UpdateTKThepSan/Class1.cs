using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace _05A_UpdateTKThepSan
{
    public class Class1
    {

        [Autodesk.AutoCAD.Runtime.CommandMethod("TKTSan")]
        public void TKThepSan()
        {
            MessageBox.Show("Đã bật được CAD");
        }
    }
}
