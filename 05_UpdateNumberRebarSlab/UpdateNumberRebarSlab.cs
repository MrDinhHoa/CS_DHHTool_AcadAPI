using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.Geometry;
using System.Windows;

namespace _05_UpdateNumberRebarSlab
{
    public class UpdateNumberRebarSlab
    {
        [CommandMethod("TKTSan")]
        public static void TKTSan()
        {
            #region Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            MessageBox.Show("Đã bật được CAD");
            #endregion
            try
            {
                var tvs = new TypedValue[]
                    {
                      new TypedValue((int)DxfCode.BlockName,"TSan_001p")
                    };

                var filter = new SelectionFilter(tvs);
                var psr = ed.GetSelection(filter);
                if (psr.Status != PromptStatus.OK) return;
                var ss = psr.Value;
            }

            catch { }
        }
    }
}
