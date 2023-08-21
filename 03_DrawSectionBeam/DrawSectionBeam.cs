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
            double cover = 25;
            double fillet = 15;
            using (Transaction Tx = acCurDb.TransactionManager.StartTransaction())
            {
                try
                {
                    Range lastData = activeSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                    int firstrow = 37;
                    int lastrow = lastData.Row;
                    for (int i = firstrow; i <= 50; i++)
                    {
                        //Lấy dữ liệu từ file excel
                        string nameBeam = (string)(activeSheet.Cells[i, 2] as Range).Value;
                        string localtion = (string)(activeSheet.Cells[i, 3] as Range).Value;
                        var width = (activeSheet.Cells[i, 6] as Range).Value;
                        var height = (activeSheet.Cells[i, 7] as Range).Value;
                        var top1_Num = (activeSheet.Cells[i, 17] as Range).Value;
                        var top1_Dia = (activeSheet.Cells[i, 18] as Range).Value;
                        BlockTable blockTable = Tx.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord blkTableRecord = Tx.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        
                        //Chèn khung ký hiệu 
                        ObjectId blockId = blockTable["BeamAll"];
                        BlockReference blockReference = new BlockReference(new Point3d(point2D.X + (i - firstrow) * 1000, point2D.Y, 0),blockId);
                        blkTableRecord.AppendEntity(blockReference);
                        Tx.AddNewlyCreatedDBObject(blockReference, true);
                        
                        //Vẽ biên dạng dầm
                        Point2d point2D1 = new Point2d(point2D.X + (i-firstrow)*1000 + 500-width/2,point2D.Y+1150-height/2);
                        Polyline p1 = Library.drawRectangle(point2D1, width,height);
                        p1.SetDatabaseDefaults();
                        blkTableRecord.AppendEntity(p1);
                        Tx.AddNewlyCreatedDBObject(p1, true);
                        p1.Closed = true;
                        p1.Layer = "S-Border";
                        p1.ColorIndex = 3;
                        
                        //Vẽ thép đai
                        Polyline stirrup = Library.drawstirrup(point2D1, width, height, fillet, cover);
                        stirrup.SetDatabaseDefaults();
                        stirrup.Closed = true;
                        stirrup.Layer = "S-Stif";
                        stirrup.ColorIndex = 171;
                        blkTableRecord.AppendEntity(stirrup);
                        Tx.AddNewlyCreatedDBObject(stirrup, true);

                        Point3d pointDim1 = new Point3d(point2D1.X, point2D1.Y + height, 0);
                        // Vẽ Kích thước ngang
                        Point3d pointDimWidth = new Point3d(point2D1.X + width, point2D1.Y + height, 0);
                        Point3d pointDimWidthlocation = new Point3d(point2D1.X + width, point2D1.Y +height+ 100, 0);
                        RotatedDimension widthdim = new RotatedDimension(0,pointDim1, pointDimWidth, pointDimWidthlocation,"", acCurDb.Dimstyle);
                        widthdim.TransformBy(ed.CurrentUserCoordinateSystem);
                        blkTableRecord.AppendEntity(widthdim);
                        Tx.AddNewlyCreatedDBObject(widthdim, true);

                        // Vẽ Kích thước dọc
                        Point3d pointDimHeight = new Point3d(point2D1.X, point2D1.Y, 0);
                        Point3d pointDimHeightlocation = new Point3d(point2D1.X -100, point2D1.Y , 0);
                        RotatedDimension heightdim = new RotatedDimension(Math.PI/2, pointDim1, pointDimHeight, pointDimHeightlocation, "", acCurDb.Dimstyle);
                        heightdim.TransformBy(ed.CurrentUserCoordinateSystem);
                        blkTableRecord.AppendEntity(heightdim);
                        Tx.AddNewlyCreatedDBObject(heightdim, true);

                        // Lấy block thép chủ
                        ObjectId tiebarId = blockTable["TieBar"];

                        // Vẽ thép chủ lớp 1
                        double fullLayer1 = width - 2 * (cover + fillet);
                        double disLayer1 = fullLayer1 / (top1_Num-1);
                        for( int j = 1; j< top1_Num + 1;)
                        {
                            double tieBarInsY = 0;
                            if (localtion.Contains("GỐI") ||localtion.Contains("END") == true)
                            { tieBarInsY = point2D1.Y + height - cover - fillet;}
                            else if (localtion.Contains("NHỊP") || localtion.Contains("SPAN") == true)
                            { tieBarInsY = point2D1.Y + cover + fillet; }    
                            BlockReference layer1Bar = new BlockReference(new Point3d(point2D1.X + cover + fillet + disLayer1*(j-1), tieBarInsY , 0), tiebarId);
                            blkTableRecord.AppendEntity(layer1Bar);
                            Tx.AddNewlyCreatedDBObject(layer1Bar, true);
                            j++;
                        }    


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
