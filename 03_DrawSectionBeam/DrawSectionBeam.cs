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

namespace _03_DrawSectionBeam
{
    public class DrawSectionBeam
    {

        [CommandMethod("drawSectionbeam")]
        [Obsolete]
        public static void drawSectionbeam()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            PromptKeywordOptions excelOption = new PromptKeywordOptions("Select/All: ");
            excelOption.Keywords.Add("Select");
            excelOption.Keywords.Add("All");
            excelOption.AllowNone = false;
            PromptResult result = ed.GetKeywords(excelOption);
        }
    }
}
