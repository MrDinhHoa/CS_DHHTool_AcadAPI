using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.Geometry;
using System.Windows;

namespace _04_SummaryFootingRebar
{
    public class SummaryFootingRebar
    {
        [CommandMethod("TKThepMong")]
        [Obsolete]

        public static void ThongKeThep()
        {
            #region Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            #endregion

            #region Get Input from user
            PromptStringOptions pStFootSize = new PromptStringOptions("\nKích thước móng (BxHxW):");
            PromptResult pStFootSizeResult = ed.GetString(pStFootSize);
            string footSizeStr = pStFootSizeResult.StringResult;
            MessageBox.Show(pStFootSizeResult.ToString());
            #endregion

        }
    }
}
