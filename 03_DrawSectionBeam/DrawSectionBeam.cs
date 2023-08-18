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
            PromptPointOptions ppo = new PromptPointOptions("Select Insert Point: ");
            PromptPointResult ppR = ed.GetPoint(ppo);
            Point3d insertPoint3D = ppR.Value;
            Point2d point2D = new Point2d(insertPoint3D.X,insertPoint3D.Y);

            using (Transaction Tx = acCurDb.TransactionManager.StartTransaction())
            {
                try
                {
                    Range lastData = activeSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                    int firstrow = 37;
                    int lastrow = lastData.Row;
                    for (int i = firstrow; i <= 90; i++)
                    {
                        string nameBeam = (string)(activeSheet.Cells[i, 2] as Range).Value;
                        string localtion = (string)(activeSheet.Cells[i, 3] as Range).Value;
                        var width = (activeSheet.Cells[i, 6] as Range).Value;
                        var height = (activeSheet.Cells[i, 7] as Range).Value;
                        var top1_Num = (activeSheet.Cells[i, 17] as Range).Value;
                        var top1_Dia = (activeSheet.Cells[i, 18] as Range).Value;
                        BlockTable blockTable = Tx.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord blkTableRecord = Tx.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        Point2d point2D1 = new Point2d(point2D.X + (i-firstrow)*1000,0);
                        //Specify the Polyline 's coordinates
                        Polyline p1 = new Polyline();
                        p1.AddVertexAt(0, new Point2d(point2D1.X, 0), 0, 0, 0);
                        p1.AddVertexAt(1, new Point2d(point2D1.X + width, 0), 0, 0, 0);
                        p1.AddVertexAt(2, new Point2d(point2D1.X + width, height), 0, 0, 0);
                        p1.AddVertexAt(3, new Point2d(point2D1.X, height), 0, 0, 0);
                        p1.AddVertexAt(4, new Point2d(point2D1.X, 0), 0, 0, 0);
                        p1.SetDatabaseDefaults();
                        blkTableRecord.AppendEntity(p1);
                        Tx.AddNewlyCreatedDBObject(p1, true);
                        p1.Closed = true;
                       
                    }
                    Tx.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("Error: " + ex.Message);
                    Tx.Abort();
                }
            }
        }
    }
}
