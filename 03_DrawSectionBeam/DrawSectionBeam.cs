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

        [CommandMethod("MCNgang")]
        [Obsolete]
        public static void drawSectionbeam()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            //PromptKeywordOptions excelOption = new PromptKeywordOptions("Select/All: ");
            //excelOption.Keywords.Add("Select");
            //excelOption.Keywords.Add("All");
            //excelOption.AllowNone = false;
            PromptIntegerOptions promptDouble = new PromptIntegerOptions("Nhập dòng cuối cùng: ");
            PromptIntegerResult lastRowNumber = ed.GetInteger(promptDouble);
            
            //PromptResult result = ed.GetKeywords(excelOption);
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
                    int lastrow = lastRowNumber.Value; 
                    for (int i = firstrow; i <= lastrow; i++)
                    {
                        #region Lấy dữ liệu từ file excel
                        string nameBeam = (string)(activeSheet.Cells[i, 2] as Range).Value;
                        string localtion = (string)(activeSheet.Cells[i, 3] as Range).Value;
                        var width = (activeSheet.Cells[i, 6] as Range).Value;
                        var height = (activeSheet.Cells[i, 7] as Range).Value;
                        var main1_Num1 = (activeSheet.Cells[i, 17] as Range).Value;
                        var main1_Dia1 = (activeSheet.Cells[i, 18] as Range).Value;
                        var main1_Num2 = (activeSheet.Cells[i, 19] as Range).Value;
                        var main1_Dia2 = (activeSheet.Cells[i, 20] as Range).Value;
                        double main1_Num = 0;
                        if(main1_Num2 != null)
                        {main1_Num = main1_Num1 + main1_Num2;}
                        else {main1_Num = main1_Num1;}
                        var main2_Num1 = (activeSheet.Cells[i, 21] as Range).Value;
                        var main2_Dia1 = (activeSheet.Cells[i, 22] as Range).Value;
                        var main2_Num2 = (activeSheet.Cells[i, 23] as Range).Value;
                        var main2_Dia2 = (activeSheet.Cells[i, 24] as Range).Value;
                        double main2_Num = 0;
                        if(main2_Num1 !=null)
                        {
                            if (main2_Num2 != null)
                            { main2_Num = main2_Num1 + main2_Num2; }
                            else { main2_Num = main2_Num1; }
                        }    
                        else if (main2_Num1 == null)
                        {
                            main2_Num = 0;
                        }    
                        var stirrup_Dia = (activeSheet.Cells[i, 8] as Range).Value;
                        var stirrup_Num = (activeSheet.Cells[i, 9] as Range).Value;
                        var stirrup_Dis = (activeSheet.Cells[i, 10] as Range).Value;
                        BlockTable blockTable = Tx.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord blkTableRecord = Tx.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        #endregion
                        #region Chèn khung ký hiệu 
                        ObjectId blockIdAll = blockTable["BeamAll"];
                        ObjectId blockIdEndSpan = blockTable["BeamEndSpan"];
                        if (localtion.Contains("TẤT CẢ" )|| localtion.Contains("ALL"))
                        {
                            BlockReference blockReference = new BlockReference(new Point3d(point2D.X + (i - firstrow) * 1000, point2D.Y, 0), blockIdAll);
                            blkTableRecord.AppendEntity(blockReference);
                            Tx.AddNewlyCreatedDBObject(blockReference, true);
                        }
                        else if (localtion == "GỐI" || localtion == "END")
                        {
                            BlockReference blockReference = new BlockReference(new Point3d(point2D.X + (i - firstrow) * 1000, point2D.Y, 0), blockIdEndSpan);
                            blkTableRecord.AppendEntity(blockReference);
                            Tx.AddNewlyCreatedDBObject(blockReference, true);
                        }
                        #endregion
                        #region Vẽ biên dạng dầm
                        Point2d point2D1 = new Point2d(point2D.X + (i-firstrow)*1000 + 500-width/2,point2D.Y+1150-height/2);
                        Polyline p1 = Library.drawRectangle(point2D1, width,height);
                        p1.SetDatabaseDefaults();
                        blkTableRecord.AppendEntity(p1);
                        Tx.AddNewlyCreatedDBObject(p1, true);
                        p1.Closed = true;
                        p1.Layer = "S-Border";
                        p1.ColorIndex = 3;
                        #endregion
                        #region Vẽ thép đai
                        Polyline stirrup = Library.drawstirrup(point2D1, width, height, fillet, cover);
                        stirrup.SetDatabaseDefaults();
                        stirrup.Closed = true;
                        stirrup.Layer = "S-Stif";
                        stirrup.ColorIndex = 171;
                        blkTableRecord.AppendEntity(stirrup);
                        Tx.AddNewlyCreatedDBObject(stirrup, true);
                        #endregion
                        #region Điểm chèn
                        Point3d pointDim1 = new Point3d(point2D1.X, point2D1.Y + height, 0);
                        #endregion
                        #region Vẽ Kích thước ngang
                        Point3d pointDimWidth = new Point3d(point2D1.X + width, point2D1.Y + height, 0);
                        Point3d pointDimWidthlocation = new Point3d(point2D1.X + width, point2D1.Y +height+ 100, 0);
                        RotatedDimension widthdim = new RotatedDimension(0,pointDim1, pointDimWidth, pointDimWidthlocation,"", acCurDb.Dimstyle);
                        widthdim.TransformBy(ed.CurrentUserCoordinateSystem);
                        blkTableRecord.AppendEntity(widthdim);
                        Tx.AddNewlyCreatedDBObject(widthdim, true);
                        #endregion
                        #region Vẽ kích thước dọc
                        Point3d pointDimHeight = new Point3d(point2D1.X, point2D1.Y, 0);
                        Point3d pointDimHeightlocation = new Point3d(point2D1.X -100, point2D1.Y , 0);
                        RotatedDimension heightdim = new RotatedDimension(Math.PI/2, pointDim1, pointDimHeight, pointDimHeightlocation, "", acCurDb.Dimstyle);
                        heightdim.TransformBy(ed.CurrentUserCoordinateSystem);
                        blkTableRecord.AppendEntity(heightdim);
                        Tx.AddNewlyCreatedDBObject(heightdim, true);
                        #endregion
                        #region Lấy block thép chủ
                        ObjectId tiebarId = blockTable["TieBar"];
                        #endregion
                        #region Vẽ thép chủ lớp 1 - Chịu kéo
                        double fullLayer1Main = width - 2 * (cover + fillet);
                        double disLayer1Main = fullLayer1Main / (main1_Num-1);
                        for( int j = 1; j< main1_Num + 1;)
                        {
                            double tieBarMainY1 = 0;
                            if (localtion.Contains("GỐI") ||localtion.Contains("END") == true)
                            { tieBarMainY1 = point2D1.Y + height - cover - fillet;}
                            else if (localtion.Contains("NHỊP") || localtion.Contains("SPAN") == true)
                            { tieBarMainY1 = point2D1.Y + cover + fillet; }    
                            BlockReference layer1Bar = new BlockReference(new Point3d(point2D1.X + cover + fillet + disLayer1Main*(j-1), tieBarMainY1, 0), tiebarId);
                            blkTableRecord.AppendEntity(layer1Bar);
                            Tx.AddNewlyCreatedDBObject(layer1Bar, true);
                            j++;
                        }
                        #endregion
                        #region Vẽ thép chủ lớp 2 - Chịu kéo
                        double fullLayer2Main = width - 2 * (cover + fillet);
                        if(main2_Num > 2)
                        {
                            double disLayer2Main = fullLayer2Main / (main2_Num - 1);
                            for (int j = 1; j < main2_Num + 1;)
                            {
                                double tieBarInsMain2 = 0;
                                if (localtion.Contains("GỐI") || localtion.Contains("END") == true)
                                { tieBarInsMain2 = point2D1.Y + height - cover - fillet - 50; }
                                else if (localtion.Contains("NHỊP") || localtion.Contains("SPAN") == true)
                                { tieBarInsMain2 = point2D1.Y + cover + fillet + 50; }
                                BlockReference layer2Bar = new BlockReference(new Point3d(point2D1.X + cover + fillet + disLayer2Main * (j - 1), tieBarInsMain2, 0), tiebarId);
                                blkTableRecord.AppendEntity(layer2Bar);
                                Tx.AddNewlyCreatedDBObject(layer2Bar, true);
                                j++;
                            }
                            Polyline hookRebarforMain2 = new Polyline();
                            if(localtion.Contains("GỐI") || localtion.Contains("END"))
                            {
                                hookRebarforMain2 = Library.drawhookRebar(new Point2d(point2D1.X + 65, point2D1.Y + height - cover - fillet - 50 + 17.5), width, cover);
                            }
                            else if (localtion.Contains("NHỊP") || localtion.Contains("SPAN"))
                            {
                                hookRebarforMain2 = Library.drawhookRebar(new Point2d(point2D1.X + 65, point2D1.Y + cover + fillet + 50 + 17.5), width, cover);
                            }
                            hookRebarforMain2.SetDatabaseDefaults();
                            hookRebarforMain2.Closed = false;
                            hookRebarforMain2.Layer = "S-Stif";
                            hookRebarforMain2.ColorIndex = 171;
                            blkTableRecord.AppendEntity(hookRebarforMain2);
                            Tx.AddNewlyCreatedDBObject(hookRebarforMain2, true);
                        }    
                        #endregion
                        #region Vẽ thép chủ - chịu nén
                        double sub_Num = 0;

                        if (localtion.Contains("ALL") || localtion.Contains("TẤT CẢ"))
                        {
                            sub_Num = (activeSheet.Cells[i, 17] as Range).Value;
                        }
                        else if (localtion == "GỐI" || localtion == "END")
                        {
                            sub_Num = (activeSheet.Cells[i + 1, 17] as Range).Value;
                        }
                        else if (localtion == "NHỊP" || localtion == "SPAN")
                        {
                            sub_Num = (activeSheet.Cells[i - 1, 17] as Range).Value;
                        }
                        double fullLayerSub = width - 2 * (cover + fillet);
                        double disLayerSub = fullLayerSub / (sub_Num - 1);
                        for (int j = 1; j < sub_Num + 1;)
                        {
                            double tieBarInsSub = 0;
                            if (localtion.Contains("GỐI") || localtion.Contains("END") == true)
                            { tieBarInsSub = point2D1.Y + cover + fillet; }
                            else if (localtion.Contains("NHỊP") || localtion.Contains("SPAN") == true)
                            { tieBarInsSub = point2D1.Y + height - cover - fillet; }
                            BlockReference layer2Bar = new BlockReference(new Point3d(point2D1.X + cover + fillet + disLayerSub * (j - 1), tieBarInsSub, 0), tiebarId);
                            blkTableRecord.AppendEntity(layer2Bar);
                            Tx.AddNewlyCreatedDBObject(layer2Bar, true);
                            j++;
                        }
                        #endregion
                        #region Vẽ thép giá + Thép móc cho thép giá
                        if(height >= 700)
                        {
                            double Add_Num = 2;
                            double fullLayerADD = width - 2 * (cover + fillet);
                            double disLayerAdd = fullLayerADD / (Add_Num - 1);
                            for (int j = 1; j < Add_Num + 1;)
                            {
                                double tieBarInsAdd = point2D1.Y + height/2;
                                BlockReference layer2Bar = new BlockReference(new Point3d(point2D1.X + cover + fillet + disLayerAdd * (j - 1), tieBarInsAdd, 0), tiebarId);
                                blkTableRecord.AppendEntity(layer2Bar);
                                Tx.AddNewlyCreatedDBObject(layer2Bar, true);
                                j++;
                            }
                            Polyline hookRebar = Library.drawhookRebar(new Point2d(point2D1.X + 65, point2D1.Y + height / 2 + 17.5), width, cover);
                            hookRebar.SetDatabaseDefaults();
                            hookRebar.Closed = false;
                            hookRebar.Layer = "S-Stif";
                            hookRebar.ColorIndex = 171;
                            blkTableRecord.AppendEntity(hookRebar);
                            Tx.AddNewlyCreatedDBObject(hookRebar, true);

                            Point3d LocationPoint = new Point3d();
                            LocationPoint = new Point3d(point2D.X + (i - firstrow) * 1000 + 500, point2D.Y + 50, 0);
                            using (DBText addRebarText = new DBText())
                            {
                                addRebarText.Position = LocationPoint;
                                addRebarText.VerticalMode = TextVerticalMode.TextVerticalMid;
                                addRebarText.HorizontalMode = TextHorizontalMode.TextCenter;
                                addRebarText.AlignmentPoint = LocationPoint;
                                addRebarText.Height = 50;
                                addRebarText.WidthFactor = 0.8;
                                addRebarText.Layer = "S-25Text-VN";
                                addRebarText.ColorIndex = 1;
                                addRebarText.TextString = "2T12 - T10a500";
                                blkTableRecord.AppendEntity(addRebarText);
                                Tx.AddNewlyCreatedDBObject(addRebarText, true);
                            }
                        }

                        #endregion
                        #region Ghi chú thép chính
                        string MainRebar1A = main1_Num1.ToString() + "T" + main1_Dia1.ToString();
                        string MainRebar1B;
                        if(main1_Num2 == null) {MainRebar1B = "";}
                        else{MainRebar1B = " + " + main1_Num2.ToString() + "T" + main1_Dia2.ToString();}    
                        string MainRebar2A;
                        if (main2_Num1 == null) { MainRebar2A = ""; }
                        else { MainRebar2A = " + " + main2_Num1.ToString() + "T" + main2_Dia1.ToString(); }
                        string MainRebar2B;
                        if (main2_Num2 == null) { MainRebar2B = ""; }
                        else { MainRebar2B = " + " + main2_Num2.ToString() + "T" + main2_Dia2.ToString(); }
                        string MainRebar = MainRebar1A + MainRebar1B + MainRebar2A + MainRebar2B;
                        Point3d alingPoint = new Point3d();
                        if (localtion.Contains("GỐI") || localtion.Contains("END") == true)
                        { alingPoint = new Point3d(point2D.X + (i - firstrow) * 1000 + 500, point2D.Y + 350, 0); }
                        else if (localtion.Contains("NHỊP") || localtion.Contains("SPAN") == true)
                        { alingPoint = new Point3d(point2D.X + (i - firstrow) * 1000 + 500, point2D.Y + 250, 0); }
                        using (DBText acText = new DBText())
                        {
                            acText.Position = alingPoint;
                            acText.VerticalMode = TextVerticalMode.TextVerticalMid;
                            acText.HorizontalMode = TextHorizontalMode.TextCenter;
                            acText.AlignmentPoint = alingPoint;
                            acText.Height = 50;
                            acText.WidthFactor = 0.8;
                            acText.Layer = "S-25Text-VN";
                            acText.ColorIndex = 1;
                            acText.TextString = MainRebar;
                            blkTableRecord.AppendEntity(acText);
                            Tx.AddNewlyCreatedDBObject(acText, true);

                        }
                        #endregion
                        #region Ghi chú thép phụ
                        double sub_Dia = 0;
                        if (localtion.Contains("ALL") || localtion.Contains("TẤT CẢ"))
                        {
                            sub_Dia = (activeSheet.Cells[i, 18] as Range).Value;
                        }
                        else if (localtion == "GỐI" || localtion == "END")
                        {
                            sub_Dia = (activeSheet.Cells[i + 1, 18] as Range).Value;
                        }
                        else if (localtion == "NHỊP" || localtion == "SPAN")
                        {
                            sub_Dia = (activeSheet.Cells[i - 1, 18] as Range).Value;
                        }
                        string SubRebar = sub_Num + "T" + sub_Dia;
                        Point3d alingPoint2 = new Point3d();
                        if (localtion.Contains("GỐI") || localtion.Contains("END") == true)
                        { alingPoint2 = new Point3d(point2D.X + (i - firstrow) * 1000 + 500, point2D.Y + 250, 0); }
                        else if (localtion.Contains("NHỊP") || localtion.Contains("SPAN") == true)
                        { alingPoint2 = new Point3d(point2D.X + (i - firstrow) * 1000 + 500, point2D.Y + 350, 0); }
                        using (DBText acText2 = new DBText())
                        {
                            acText2.Position = alingPoint2;
                            acText2.VerticalMode = TextVerticalMode.TextVerticalMid;
                            acText2.HorizontalMode = TextHorizontalMode.TextCenter;
                            acText2.AlignmentPoint = alingPoint2;
                            acText2.Height = 50;
                            acText2.WidthFactor = 0.8;
                            acText2.Layer = "S-25Text-VN";
                            acText2.ColorIndex = 1;
                            acText2.TextString = SubRebar;
                            blkTableRecord.AppendEntity(acText2);
                            Tx.AddNewlyCreatedDBObject(acText2, true);
                        }
                        #endregion
                        #region Ghi chú tên dầm
                        Point3d alingPoint3 = new Point3d();
                        if (localtion.Contains("ALL") || localtion.Contains("TẤT CẢ") == true)
                        { alingPoint3 = new Point3d(point2D.X + (i - firstrow) * 1000 + 500, point2D.Y + 2100, 0); }
                        else if (localtion == "GỐI" || localtion == "END" )
                        { alingPoint3 = new Point3d(point2D.X + (i - firstrow) * 1000 + 1000, point2D.Y + 2100, 0); }
                        if(nameBeam != null)
                        {
                            using (DBText acText3 = new DBText())
                            {
                                acText3.Position = alingPoint3;
                                acText3.VerticalMode = TextVerticalMode.TextVerticalMid;
                                acText3.HorizontalMode = TextHorizontalMode.TextCenter;
                                acText3.AlignmentPoint = alingPoint3;
                                acText3.Height = 100;
                                acText3.WidthFactor = 0.8;
                                acText3.Layer = "S-25Text-VN";
                                acText3.ColorIndex = 1;
                                acText3.TextString = nameBeam;
                                blkTableRecord.AppendEntity(acText3);
                                Tx.AddNewlyCreatedDBObject(acText3, true);
                            }
                        }

                        #endregion
                        #region Ghi chú thép đai
                        string stirrupRebar;
                        if(stirrup_Dia < 9) {stirrupRebar = "R" + stirrup_Dia + "a" + stirrup_Dis;}  
                        else { stirrupRebar = "T" + stirrup_Dia + "a" + stirrup_Dis; }
                        Point3d LocationPointStirrup = new Point3d();
                        LocationPointStirrup = new Point3d(point2D.X + (i - firstrow) * 1000 + 500, point2D.Y + 150, 0);
                        using (DBText stirRebarText = new DBText())
                        {
                            stirRebarText.Position = LocationPointStirrup;
                            stirRebarText.VerticalMode = TextVerticalMode.TextVerticalMid;
                            stirRebarText.HorizontalMode = TextHorizontalMode.TextCenter;
                            stirRebarText.AlignmentPoint = LocationPointStirrup;
                            stirRebarText.Height = 50;
                            stirRebarText.WidthFactor = 0.8;
                            stirRebarText.Layer = "S-25Text-VN";
                            stirRebarText.ColorIndex = 1;
                            stirRebarText.TextString = stirrupRebar;
                            blkTableRecord.AppendEntity(stirRebarText);
                            Tx.AddNewlyCreatedDBObject(stirRebarText, true);
                        }
                        #endregion
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
