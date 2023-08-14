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
            Excel.Application oExcelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Workbook activeWorkbook = oExcelApp.ActiveWorkbook;
            Worksheet activeSheet = activeWorkbook.ActiveSheet;
            Range lastData = activeSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            int firstrow = 37;
            int lastrow = lastData.Row;
            for(int i = firstrow; i<= lastrow; i++)
            {
                string nameBeam = (string)(activeSheet.Cells[i, 2] as Range).Value;
                string localtion = (string)(activeSheet.Cells[i,3] as Range).Value;
            }    
                
            using (Transaction Tx = acCurDb.TransactionManager.StartTransaction())
            {
                BlockTableRecord bt = Tx.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                foreach (ObjectId id in bt)
                {
                    if (id.ObjectClass.Name != "AcDbOle2Frame")
                        continue;

                    Ole2Frame oleFrame = Tx.GetObject(id, OpenMode.ForRead) as Ole2Frame;

                    if (!oleFrame.IsLinked)
                    {
                        Workbook wb = oleFrame.OleObject as Workbook;

                        Microsoft.Office.Interop.Excel.Worksheet ws = wb.ActiveSheet;

                        Microsoft.Office.Interop.Excel.Range range = ws.UsedRange;
                        for (int row = 1; row <= 15; row++)
                        {
                            for (int col = 1; col <= range.Columns.Count; col++)
                            {
                                ed.WriteMessage(String.Format("{0}{1}{2}-{3}", Environment.NewLine, row, col, Convert.ToString((range.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Value2)));
                            }
                        }
                    }
                }
            }
        }
    }
}
