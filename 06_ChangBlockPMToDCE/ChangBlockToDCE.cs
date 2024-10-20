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
namespace _06_ChangBlockPMToDCE
{
    public class ChangBlockToDCE
    {
        [CommandMethod("UpToDCEBlock")]
        [Obsolete]
        public static void UpToDCEBlock()
        {
            #region Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;

            #endregion

            #region 
            try
            {
                // Set tỉ lệ
                PromptDoubleOptions promptDouble = new PromptDoubleOptions("\nChọn Tỉ lệ: ");
                promptDouble.DefaultValue = 100;
                PromptDoubleResult lastRowNumber = ed.GetDouble(promptDouble);
                double TL = lastRowNumber.Value;
                // Chọn điểm chèn
                PromptPointResult pPtRes;
                PromptPointOptions pPtOpts = new PromptPointOptions("\nChọn điểm chèn: ");
                pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                Point3d ptStart = pPtRes.Value;

                //Default Value
                int CountWhile = 0;
                int CountBlock = 0;
                while (true)
                {
                    PromptSelectionResult psr = ed.GetSelection();
                    if (psr.Status != PromptStatus.OK) break;
                    SelectionSet ss = psr.Value;
                    using (Transaction tr = acDoc.TransactionManager.StartTransaction())
                    {
                        List<BlockRebarDCEInfor> listRebarInfor = new List<BlockRebarDCEInfor>();
                        if (ss.Count > 0)
                        {
                            foreach (SelectedObject o in ss)
                            {
                                BlockRebarDCEInfor blockRebarInfor = new BlockRebarDCEInfor();
                                ObjectId objectId = o.ObjectId;
                                using (BlockReference oBlock = objectId.Open(OpenMode.ForRead) as BlockReference)
                                {
                                    if (oBlock.IsDynamicBlock)
                                    {
                                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(oBlock.DynamicBlockTableRecord, OpenMode.ForRead);
                                        blockRebarInfor.BlockName = btr.Name;
                                    }
                                    else { blockRebarInfor.BlockName = oBlock.Name; }
                                    #region Lấy các Attribute giống nhau
                                    var attSample = oBlock.AttributeCollection;
                                    foreach (ObjectId objectId1 in attSample)
                                    {
                                        using (AttributeReference attRef = objectId1.Open(OpenMode.ForRead) as AttributeReference)
                                        {
                                            string tag = attRef.Tag;
                                            switch (tag)
                                            {
                                                case "NO":
                                                    blockRebarInfor.SH = attRef.TextString;
                                                    break;
                                                // Xu ly so lieu
                                                case "DIA":
                                                    blockRebarInfor.DK = attRef.TextString;
                                                    break;
                                                case "NIE":
                                                    blockRebarInfor.SL = attRef.TextString;
                                                    break;
                                                case "QOE":
                                                    blockRebarInfor.SCK = attRef.TextString;
                                                    break;
                                            }

                                        }
                                    }
                                    #endregion
                                    #region Lấy các Attribute khác nhau
                                    var attDiffer = oBlock.AttributeCollection;

                                    switch (blockRebarInfor.BlockName)
                                    {
                                        case "TKCot_01a":
                                            foreach (ObjectId objectId1 in attDiffer)
                                            {
                                                using (AttributeReference attRef = objectId1.Open(OpenMode.ForRead) as AttributeReference)
                                                {
                                                    string tag = attRef.Tag;
                                                    switch (tag)
                                                    {
                                                        case "D1":
                                                            blockRebarInfor.L = attRef.TextString;
                                                            break;
                                                    }

                                                }
                                            }
                                            break;
                                        case "TKCot_01c":
                                            foreach (ObjectId objectId1 in attSample)
                                            {
                                                using (AttributeReference attRef = objectId1.Open(OpenMode.ForRead) as AttributeReference)
                                                {
                                                    string tag = attRef.Tag;
                                                    switch (tag)
                                                    {
                                                        case "D2":
                                                            blockRebarInfor.L1 = attRef.TextString;
                                                            break;
                                                        case "D1":
                                                            blockRebarInfor.L2 = attRef.TextString;
                                                            break;
                                                        case "D3":
                                                            blockRebarInfor.L3 = attRef.TextString;
                                                            break;
                                                    }

                                                }
                                            }
                                            break;
                                        case "TKCot_01b":
                                            foreach (ObjectId objectId1 in attSample)
                                            {
                                                using (AttributeReference attRef = objectId1.Open(OpenMode.ForRead) as AttributeReference)
                                                {
                                                    string tag = attRef.Tag;
                                                    switch (tag)
                                                    {
                                                        case "D1":
                                                            blockRebarInfor.L1 = attRef.TextString;
                                                            break;
                                                        case "D2":
                                                            blockRebarInfor.L2 = attRef.TextString;
                                                            break;
                                                    }

                                                }
                                            }
                                            break;
                                        case "TKCot_03":
                                            foreach (ObjectId objectId1 in attSample)
                                            {
                                                using (AttributeReference attRef = objectId1.Open(OpenMode.ForRead) as AttributeReference)
                                                {
                                                    blockRebarInfor.M = "50";
                                                    string tag = attRef.Tag;
                                                    switch (tag)
                                                    {
                                                        case "D1":
                                                            blockRebarInfor.L1 = attRef.TextString;
                                                            break;
                                                        case "D2":
                                                            blockRebarInfor.L2 = attRef.TextString;
                                                            break;
                                                    }

                                                }
                                            }
                                            break;
                                        case "TKCot_02":
                                            foreach (ObjectId objectId1 in attSample)
                                            {
                                                using (AttributeReference attRef = objectId1.Open(OpenMode.ForRead) as AttributeReference)
                                                {
                                                    blockRebarInfor.M1 = "50";
                                                    blockRebarInfor.M2 = "50";
                                                    string tag = attRef.Tag;
                                                    switch (tag)
                                                    {
                                                        case "D1":
                                                            blockRebarInfor.L = attRef.TextString;
                                                            break;
                                                    }
                                                }
                                            }
                                            break;
                                        case "TKCot_06a":
                                            foreach (ObjectId objectId1 in attSample)
                                            {
                                                using (AttributeReference attRef = objectId1.Open(OpenMode.ForRead) as AttributeReference)
                                                {
                                                    string tag = attRef.Tag;
                                                    switch (tag)
                                                    {
                                                        case "D1":
                                                            blockRebarInfor.L1 = attRef.TextString;
                                                            break;
                                                        case "D2":
                                                            blockRebarInfor.L2 = attRef.TextString;
                                                            break;
                                                        case "D3":
                                                            blockRebarInfor.L3 = attRef.TextString;
                                                            break;
                                                        case "D4":
                                                            blockRebarInfor.L4 = attRef.TextString;
                                                            break;
                                                        case "D5":
                                                            blockRebarInfor.L5 = attRef.TextString;
                                                            break;
                                                    }
                                                }
                                            }
                                            break;
                                    }

                                    #endregion
                                }
                                listRebarInfor.Add(blockRebarInfor);
                            }
                        }

                        List<BlockRebarDCEInfor> SortListQuery = listRebarInfor.OrderBy(x => x.SH).ToList();

                        SortListQuery.Sort(new NaturalStringComparer());

                        List<BlockRebarDCEInfor> sortList = SortListQuery.ToList();
                        BlockRebarDCEInfor TitleBlock = new BlockRebarDCEInfor();
                        TitleBlock.BlockName = "TKTitle";
                        sortList.Insert(0, TitleBlock);
                        for (int i = 0; i < sortList.Count; i++)
                        {
                            string blockDCEName = null;
                            if (blockDCEName == null)
                            {
                                switch (sortList[i].BlockName)
                                {
                                    case "TKTitle":
                                        blockDCEName = "TKTitle";
                                        break;
                                    case "TKCot_01a":
                                        blockDCEName = "TK1";
                                        break;
                                    case "TKCot_01c":
                                        blockDCEName = "TK7";
                                        break;
                                    case "TKCot_01b":
                                        blockDCEName = "TK2";
                                        break;
                                    case "TKCot_02":
                                        blockDCEName = "TK8";
                                        break;
                                    case "TKCot_03":
                                        blockDCEName = "TK9";
                                        break;
                                    case "TKCot_06a":
                                        blockDCEName = "TK10";
                                        break;
                                    default:
                                        blockDCEName = "TK1";
                                        break;

                                }
                            }
                            BlockTable bt = acCurDb.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord blockDef = bt[blockDCEName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                            BlockTableRecord ms = bt[BlockTableRecord.ModelSpace].GetObject(OpenMode.ForWrite) as BlockTableRecord;
                            //Set Point Insert 
                            double pointY = 0;
                            if (CountWhile == 0)
                            {
                                if (i == 0) { pointY = 0; }
                                else if (i == 1) { pointY = 35 * TL; }
                                else { pointY = (35 + (i - 1) * 8) * TL; }
                            }
                            else
                            {
                                if (i == 0) { continue; }
                                else { pointY = (35 + CountBlock * 8 + (i - 1) * 8) * TL; }
                            }
                            Point3d point = new Point3d(ptStart.X, ptStart.Y - pointY, 0.0);

                            //Set Insert
                            if (i == 0)
                            {
                                if (CountWhile == 0)
                                {
                                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                                    {
                                        using (BlockReference blockRef = new BlockReference(point, blockDef.ObjectId))
                                        {
                                            blockRef.TransformBy(Matrix3d.Scaling(TL, point)); //Scale Block
                                            ms.AppendEntity(blockRef);
                                            acTrans.AddNewlyCreatedDBObject(blockRef, true);
                                        }
                                        acTrans.Commit();
                                    }
                                }
                                else { continue; }
                            }
                            else
                            {
                                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                                {
                                    using (BlockReference blockRef = new BlockReference(point, blockDef.ObjectId))
                                    {
                                        blockRef.TransformBy(Matrix3d.Scaling(TL, point)); //Scale Block
                                        ms.AppendEntity(blockRef);
                                        acTrans.AddNewlyCreatedDBObject(blockRef, true);
                                        #region Atribute
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
                                                    // Gán các attribute 
                                                    switch (Tag)
                                                    {
                                                        case "SH": //Case kich thuoc chinh
                                                            attRef.TextString = sortList[i].SH;
                                                            break;
                                                        case "DK": //Case kich thuoc chinh
                                                            attRef.TextString = sortList[i].DK;
                                                            break;
                                                        case "SL": //Case kich thuoc chinh
                                                            attRef.TextString = sortList[i].SL;
                                                            break;
                                                        case "SKC": //Case kich thuoc chinh
                                                            attRef.TextString = sortList[i].SCK;
                                                            break;

                                                    }
                                                    switch (blockDCEName)
                                                    {
                                                        case "TK1":
                                                            {
                                                                switch (Tag)
                                                                {
                                                                    case "L": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L;
                                                                        break;
                                                                }
                                                            }
                                                            break;
                                                        case "TK7":
                                                            {
                                                                switch (Tag)
                                                                {
                                                                    case "L1": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L1;
                                                                        break;
                                                                    case "L2": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L2;
                                                                        break;
                                                                    case "L3": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L3;
                                                                        break;
                                                                }
                                                            }
                                                            break;
                                                        case "TK2":
                                                            {
                                                                switch (Tag)
                                                                {
                                                                    case "L1": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L1;
                                                                        break;
                                                                    case "L2": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L2;
                                                                        break;
                                                                }
                                                            }
                                                            break;
                                                        case "TK9":
                                                            {
                                                                switch (Tag)
                                                                {
                                                                    case "L1": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L1;
                                                                        break;
                                                                    case "L2": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L2;
                                                                        break;
                                                                    case "M": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].M;
                                                                        break;
                                                                }
                                                            }
                                                            break;
                                                        case "TK8":
                                                            {
                                                                switch (Tag)
                                                                {
                                                                    case "M1": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].M1;
                                                                        break;
                                                                    case "M2": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].M2;
                                                                        break;
                                                                    case "L": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L;
                                                                        break;
                                                                }
                                                            }
                                                            break;
                                                        case "TK10":
                                                            {
                                                                switch (Tag)
                                                                {
                                                                    case "L1": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L1;
                                                                        break;
                                                                    case "L2": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L2;
                                                                        break;
                                                                    case "L3": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L3;
                                                                        break;
                                                                    case "L4": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L4;
                                                                        break;
                                                                    case "L5": //Case kich thuoc chinh
                                                                        attRef.TextString = sortList[i].L5;
                                                                        break;
                                                                }
                                                            }
                                                            break;
                                                    }
                                                    blockRef.AttributeCollection.AppendAttribute(attRef);
                                                    tr.AddNewlyCreatedDBObject(attRef, true);
                                                }

                                            }
                                        }
                                        #endregion
                                    }
                                    acTrans.Commit();
                                }
                            }
                        }
                        CountBlock = CountBlock + ss.Count;
                        tr.Commit();
                    }
                    CountWhile++;
                }
            }
            catch { }

            #endregion

        }
    }
}
