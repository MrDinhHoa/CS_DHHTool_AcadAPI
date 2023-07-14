using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Excel = Microsoft.Office.Interop.Excel;
using Autodesk.AutoCAD.Windows;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Microsoft.Office.Interop.Excel;
using Autodesk.AutoCAD.EditorInput;

namespace _02_TextByCoordinate
{
    public static class Extension
    {
        /// <summary>
        /// Strips a Point3d down to a Point2d by simply ignoring the Z ordinate.
        /// </summary>
        /// <param name="pt">The Point3d to strip.</param>
        /// <returns>The stripped Point2d.</returns>
        public static Point2d Strip(this Point3d pt)
        { return new Point2d(pt.X, pt.Y); }
    }

    public class TextByCoordinate
    {
        [CommandMethod("NameForPile")]
        [Obsolete]
        public static void NameForPile()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            PromptStringOptions Prefix = new PromptStringOptions("Type Prefix: ");
            PromptResult prefixResult = ed.GetString(Prefix);

            PromptKeywordOptions direction = new PromptKeywordOptions("Direction");
            direction.Keywords.Add("XbyY");
            direction.Keywords.Add("Ybyx");
            direction.AllowNone = false;

            PromptResult result = ed.GetKeywords(direction);


            //Chọn block 
            TypedValue[] tvs = new TypedValue[]
               {new TypedValue((int)DxfCode.Start, "INSERT") };
            SelectionFilter filter = new SelectionFilter(tvs);
            PromptSelectionResult psr = ed.GetSelection(filter);
            SelectionSet ss = psr.Value;
            List<Point2d> point2Ds = new List<Point2d>();
            for (int i = 0; i < ss.Count; i++)
            {
                ObjectId objID = ss[i].ObjectId;
                using (BlockReference oBlock = objID.Open(OpenMode.ForRead) as BlockReference)
                {
                    Point2d insertPoint = oBlock.Position.Strip();
                    point2Ds.Add(insertPoint);
                }
            }

            List<Point2d> listPoint2DSort = new List<Point2d>();


            if(result.StringResult == "XbyY")
            {
                listPoint2DSort = point2Ds.OrderBy(x => Math.Round(x.X,0))
                                .ThenBy(x => Math.Round(x.Y,0))
                                .ToList();
            }
            else
            {
                listPoint2DSort = point2Ds.OrderBy(x => Math.Round(x.Y, 0))
                                .ThenBy(x => Math.Round(x.X, 0))
                                .ToList();
            } 
                

            try
            {

                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    // Open the Block table for read
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    // Open the Block table record Model space for write
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    // Create a single-line text object
                    Polyline acPoly = new Polyline();
                    acPoly.SetDatabaseDefaults();
                    for (int i = 0; i < listPoint2DSort.Count; i++)
                    {
                        string textInsert = prefixResult.StringResult.ToString() + (i + 1).ToString();
                        DBText acText = new DBText();
                        acText.SetDatabaseDefaults();
                        acText.Position = new Point3d(listPoint2DSort[i].X, listPoint2DSort[i].Y, 0);
                        acText.Height = 500;
                        acText.TextString = textInsert;
                        acBlkTblRec.AppendEntity(acText);
                        acTrans.AddNewlyCreatedDBObject(acText, true);
                        acPoly.AddVertexAt(i, new Point2d(listPoint2DSort[i].X, listPoint2DSort[i].Y), 0, 0, 0);
                        // Add the new object to the block table record and the transaction
                    }
                    acBlkTblRec.AppendEntity(acPoly);
                    acTrans.AddNewlyCreatedDBObject(acPoly, true);
                    // Save the changes and dispose of the transaction
                    acTrans.Commit();
                }
                //    System.Windows.Forms.OpenFileDialog file = new System.Windows.Forms.OpenFileDialog();
                //{
                //    file.InitialDirectory = "C:\\";
                //    file.DefaultExt = "xls";
                //    file.Filter = "Excel File(*.xls; *.xlsx; *.xlsm)|*.xls; *.xlsx" + "|All Files (*.*)|*.*";
                //    file.FilterIndex = 1;
                //    file.RestoreDirectory = true;
                //}
                //if (file.ShowDialog() == DialogResult.OK) //if there is a file chosen by the user
                //{
                //    Excel.Application xlsApp = new Excel.Application();
                //    Workbook xlsworkbook = xlsApp.Workbooks.Open(file.FileName);
                //    Worksheet worksheet = xlsworkbook.Worksheets["Summary"];
                //    Range xlRange = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                //    int row = xlRange.Row;
                //    // Start a transaction

                //        }
                //    }


            }
            catch
            {

            }


        }
    }
}
