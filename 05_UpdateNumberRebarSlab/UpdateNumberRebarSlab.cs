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
        [Obsolete]
        public static void TKTSan()
        {
            #region Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            #endregion
            try
            {
                var psr = ed.GetSelection();
                if (psr.Status != PromptStatus.OK) return;
                var ss = psr.Value;
                List<RebarSlabInfor> listRebarInfor = new List<RebarSlabInfor>();
                if (ss.Count > 0)
                {
                    foreach (SelectedObject o in ss)
                    {
                        RebarSlabInfor rebarSlabInfor = new RebarSlabInfor();
                        rebarSlabInfor.QOE = "1";
                        ObjectId objectId = o.ObjectId;
                        using (BlockReference oBlock = objectId.Open(OpenMode.ForRead) as BlockReference)
                        {
                            var att = oBlock.AttributeCollection;
                            double Dia = 0;
                            double Distance = 0;
                            double TotalDistance = 0;
                            double LentghOne = 0;
                            double LentghAll = 0;
                            double D1 = 0;
                            double D2 = 0;
                            double D3 = 0;
                            foreach (ObjectId objectId1 in att)
                            {
                                using (AttributeReference attRef = objectId1.Open(OpenMode.ForRead) as AttributeReference)
                                {
                                    string tag = attRef.Tag;
                                    switch (tag)
                                    {
                                        case "0":
                                            rebarSlabInfor.NO = attRef.TextString;
                                            break;

                                        // Xu ly so lieu
                                        case "1":
                                            string D1string = attRef.TextString;
                                            Distance = Convert.ToDouble(D1string.Substring(D1string.IndexOf("a") + 1));
                                            Dia = Convert.ToDouble(D1string.Substring(3, D1string.IndexOf("a") - 3));
                                            rebarSlabInfor.DIA = Dia.ToString();

                                            break;
                                        case "2":
                                            string Tag2string = attRef.TextString;
                                            if (Tag2string.Contains("."))
                                            { rebarSlabInfor.D1 = attRef.TextString.Substring(0, attRef.TextString.IndexOf(".")); }
                                            else { rebarSlabInfor.D1 = attRef.TextString; }
                                            D1 = Convert.ToDouble(attRef.TextString);
                                            break;
                                        case "3":
                                            string Tag3string = attRef.TextString;
                                            if (Tag3string.Contains("."))
                                            { rebarSlabInfor.D2 = attRef.TextString.Substring(0, attRef.TextString.IndexOf(".")); }
                                            else { rebarSlabInfor.D2 = attRef.TextString; }
                                            D2 = Convert.ToDouble(attRef.TextString);
                                            break;
                                        case "4":
                                            string Tag4string = attRef.TextString;
                                            if (Tag4string.Contains("."))
                                            { rebarSlabInfor.D3 = attRef.TextString.Substring(0, attRef.TextString.IndexOf(".")); }
                                            else { rebarSlabInfor.D3 = attRef.TextString; }
                                            D3 = Convert.ToDouble(attRef.TextString);
                                            break;
                                        // Xu ly so Lieu
                                        case "5":
                                            TotalDistance = Convert.ToDouble(attRef.TextString);
                                            break;

                                    }

                                }
                            }
                            if (Distance != 0)
                            { rebarSlabInfor.NIE = (Math.Ceiling(TotalDistance / Distance)).ToString(); }
                            LentghOne = (D1 + D2 + D3) / 1000 + 40 * Dia * 0.001 * Math.Truncate((D1 + D2 + D3) / 1000 / 11.7);
                            LentghAll = LentghOne * Math.Ceiling(TotalDistance / Distance);
                            rebarSlabInfor.LO = LentghOne.ToString();
                            rebarSlabInfor.LA = LentghAll.ToString();
                        }
                        listRebarInfor.Add(rebarSlabInfor);

                    }
                }
                var SortListQuery = listRebarInfor.OrderBy(x => x.NO);
                List<RebarSlabInfor> sortList = SortListQuery.ToList();
                for (int i = 0; i < sortList.Count; i++)
                {
                    #region Transaction
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {
                        string blockName = null;

                        if ((listRebarInfor[i].D2 == "" || listRebarInfor[i].D2 == null || listRebarInfor[i].D2 == "0")
                            && ((listRebarInfor[i].D3 == "" || listRebarInfor[i].D3 == null || listRebarInfor[i].D3 == "0")))
                        { blockName = "TKCot_01a"; }
                        else if ((listRebarInfor[i].D2 == "" || listRebarInfor[i].D2 == null || listRebarInfor[i].D2 == "0"
                                 || listRebarInfor[i].D3 == "" || listRebarInfor[i].D3 == null || listRebarInfor[i].D3 == "0"))
                        { blockName = "TKCot_01b"; }
                        else { blockName = "TKCot_01c"; }
                        BlockTable bt = acCurDb.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord blockDef = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        BlockTableRecord ms = bt[BlockTableRecord.ModelSpace].GetObject(OpenMode.ForWrite) as BlockTableRecord;
                        Point3d point = new Point3d(-200000, -i * 950, 0.0);
                        using (BlockReference blockRef = new BlockReference(point, blockDef.ObjectId))
                        {
                            ms.AppendEntity(blockRef);
                            acTrans.AddNewlyCreatedDBObject(blockRef, true);
                            // AttributeDefinitions
                            foreach (ObjectId id in blockDef)
                            {
                                DBObject obj = id.GetObject(OpenMode.ForRead);
                                AttributeDefinition attDef = obj as AttributeDefinition;
                                if ((attDef != null) && (!attDef.Constant))
                                {
                                    using (AttributeReference attRef = new AttributeReference())
                                    {
                                        string Tag = attDef.Tag;
                                        attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                                        if (blockName == "TKCot_01b")
                                        {
                                            switch (Tag)
                                            {
                                                case "D2": //Case kich thuoc chinh
                                                    attRef.TextString = listRebarInfor[i].D1;
                                                    break;
                                                case "D1": // Case kich thuoc phu 1
                                                    if (listRebarInfor[i].D2 == "0")
                                                    { attRef.TextString = listRebarInfor[i].D3; }
                                                    else
                                                    { attRef.TextString = listRebarInfor[i].D2; }
                                                    break;
                                                case "NO":
                                                    attRef.TextString = listRebarInfor[i].NO;
                                                    break;
                                                case "NIE":
                                                    attRef.TextString = listRebarInfor[i].NIE;
                                                    break;
                                                case "DIA":
                                                    attRef.TextString = listRebarInfor[i].DIA;
                                                    break;
                                                case "QOE":
                                                    attRef.TextString = listRebarInfor[i].QOE;
                                                    break;
                                                case "LO":
                                                    attRef.TextString = listRebarInfor[i].LO;
                                                    break;
                                                case "LA":
                                                    attRef.TextString = listRebarInfor[i].LA;
                                                    break;
                                            }
                                        }
                                        else
                                        {
                                            switch (Tag)
                                            {
                                                case "D2":
                                                    attRef.TextString = listRebarInfor[i].D2;
                                                    break;
                                                case "D1":
                                                    attRef.TextString = listRebarInfor[i].D1;
                                                    break;
                                                case "D3":
                                                    attRef.TextString = listRebarInfor[i].D3;
                                                    break;
                                                case "NO":
                                                    attRef.TextString = listRebarInfor[i].NO;
                                                    break;
                                                case "NIE":
                                                    attRef.TextString = listRebarInfor[i].NIE;
                                                    break;
                                                case "DIA":
                                                    attRef.TextString = listRebarInfor[i].DIA;
                                                    break;
                                                case "QOE":
                                                    attRef.TextString = listRebarInfor[i].QOE;
                                                    break;
                                                case "LO":
                                                    attRef.TextString = listRebarInfor[i].LO;
                                                    break;
                                                case "LA":
                                                    attRef.TextString = listRebarInfor[i].LA;
                                                    break;
                                            }
                                        }
                                        blockRef.AttributeCollection.AppendAttribute(attRef);
                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                    }

                                }
                            }
                        }
                        acTrans.Commit();
                    }
                    #endregion
                }
            }
            catch { }
        }
    }
}
