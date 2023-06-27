using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Exception = Autodesk.AutoCAD.Runtime.Exception;
using Autodesk.AutoCAD.DatabaseServices;

namespace _01_LayoutByXrefBlock
{
    public class LayoutByXrefBlock
    {
        [CommandMethod("LayoutbyXrefBlock")]
        public void LayoutbyXrefBlock()
        {
            var aDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = aDoc.Editor;
            try
            {
                //Create the prompts
                PromptKeywordOptions pko = new PromptKeywordOptions("Chọn Xref / Block: ");
                pko.Keywords.Add("XRef");
                pko.Keywords.Add("bLock");
                pko.AllowNone = false;

                //Get user input
                PromptResult pResult = ed.GetKeywords(pko);
                string strfilter = "";
                if(pResult.StringResult == "XRef")
                { strfilter = "XREF"; }
                else if(pResult.StringResult == "bLock")
                { strfilter = "INSERT"; }    
                var tvs = new TypedValue[]
                {new TypedValue((int)DxfCode.Start, strfilter) };

                var filter = new SelectionFilter(tvs);

                var psr = ed.GetSelection("Chọn ",filter);
            }
            catch (Exception) 
            { 
            }

            
        }
    }
}
