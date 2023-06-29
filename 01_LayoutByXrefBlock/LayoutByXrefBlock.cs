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
        [CommandMethod("LayByXrBl")]
        [Obsolete]
        public void LayoutbyXrefBlock()
        {
            var aDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = aDoc.Editor;
            try
            {
                //Chọn block mẫu
                string blockName;
                TypedValue[] tvs = new TypedValue[]
                {new TypedValue((int)DxfCode.Start, "INSERT") };
                SelectionFilter filter = new SelectionFilter(tvs);
                PromptSelectionResult psr = ed.GetSelection(filter);
                SelectionSet ss = psr.Value;
                ObjectId  objID  = ss[0].ObjectId;
                using(BlockReference oBlock = objID.Open(OpenMode.ForRead) as BlockReference)
                { blockName = oBlock.Name;}

                //Chọn tất cả các block dựa vào block mẫu
                TypedValue[] tvbyName = new TypedValue[]
                {new TypedValue((int)DxfCode.BlockName, blockName) };
                SelectionFilter filterbyName = new SelectionFilter(tvbyName);
                PromptSelectionResult psrbyName = ed.GetSelection(filterbyName);
                SelectionSet ssbyName = psrbyName.Value;
                foreach(SelectedObject v in ssbyName) 
                {
                    ObjectId objIDbyName = v.ObjectId;
                    using (BlockReference oblByName = objIDbyName.Open(OpenMode.ForRead) as BlockReference);
                }
             

            }
            catch (Exception) 
            { 
            }

            
        }
    }
}
