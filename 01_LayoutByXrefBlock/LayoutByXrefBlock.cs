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
using System.Collections;
using Autodesk.AutoCAD.Geometry;
using static System.Net.Mime.MediaTypeNames;

namespace _01_LayoutByXrefBlock
{
    public static class Extensions
    {
        /// <summary>
        /// Reverses the order of the X and Y properties of a Point2d.
        /// </summary>
        /// <param name="flip">Boolean indicating whether to reverse or not.</param>
        /// <returns>The original Point2d or the reversed version.</returns>
        public static Point2d Swap(this Point2d pt, bool flip = true)
        { return flip ? new Point2d(pt.Y, pt.X) : pt; }

        /// <summary>
        /// Pads a Point2d with a zero Z value, returning a Point3d.
        /// </summary>
        /// <param name="pt">The Point2d to pad.</param>
        /// <returns>The padded Point3d.</returns>
        public static Point3d Pad(this Point2d pt)
        { return new Point3d(pt.X, pt.Y, 0); }

        /// <summary>
        /// Strips a Point3d down to a Point2d by simply ignoring the Z ordinate.
        /// </summary>
        /// <param name="pt">The Point3d to strip.</param>
        /// <returns>The stripped Point2d.</returns>
        public static Point2d Strip(this Point3d pt)
        { return new Point2d(pt.X, pt.Y); }

        /// <summary>
        /// Creates a layout with the specified name and optionally makes it current.
        /// </summary>
        /// <param name="name">The name of the viewport.</param>
        /// <param name="select">Whether to select it.</param>
        /// <returns>The ObjectId of the newly created viewport.</returns>
        public static ObjectId CreateAndMakeLayoutCurrent(this LayoutManager lm, string name, bool select = true)
        {
            // First try to get the layout
            var id = lm.GetLayoutId(name);
            // If it doesn't exist, we create it
            if (!id.IsValid)
            { id = lm.CreateLayout(name); }
            // And finally we select it
            if (select)
            { lm.CurrentLayout = name; }
            return id;
        }

        /// <summary>
        /// Applies an action to the specified viewport from this layout.
        /// Creates a new viewport if none is found withthat number.
        /// </summary>
        /// <param name="tr">The transaction to use to open the viewports.</param>
        /// <param name="vpNum">The number of the target viewport.</param>
        /// <param name="f">The action to apply to each of the viewports.</param>
        public static void ApplyToViewport(this Layout lay, Transaction tr, int vpNum, Action<Viewport> f)
        {
            ObjectIdCollection vpIds = lay.GetViewports();
            Viewport vp = null;
            foreach (ObjectId vpId in vpIds)
            {
                Viewport vp2 = tr.GetObject(vpId, OpenMode.ForWrite) as Viewport;
                if (vp2 != null && vp2.Number == vpNum)
                {
                    // We have found our viewport, so call the action
                    vp = vp2;
                    break;
                }
            }

            if (vp == null)
            {
                // We have not found our viewport, so create one
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(lay.BlockTableRecordId, OpenMode.ForWrite);
                vp = new Viewport();
                // Add it to the database
                btr.AppendEntity(vp);
                tr.AddNewlyCreatedDBObject(vp, true);
                // Turn it - and its grid - on
                vp.On = true;
                vp.GridOn = true;
            }
            // Finally we call our function on it
            f(vp);
        }

        /// <summary>
        /// Apply plot settings to the provided layout.
        /// </summary>
        /// <param name="pageSize">The canonical media name for our page size.</param>
        /// <param name="styleSheet">The pen settings file (ctb or stb).</param>
        /// <param name="devices">The name of the output device.</param>
        public static void SetPlotSettings(this Layout lay, string pageSize, string styleSheet, string device)
        {
            using (PlotSettings ps = new PlotSettings(lay.ModelType))
            {
                ps.CopyFrom(lay);
                PlotSettingsValidator psv = PlotSettingsValidator.Current;
                // Set the device
                System.Collections.Specialized.StringCollection devs = psv.GetPlotDeviceList();
                if (devs.Contains(device))
                {
                    psv.SetPlotConfigurationName(ps, device, null);
                    psv.RefreshLists(ps);
                }
                // Set the media name/size
                System.Collections.Specialized.StringCollection mns = psv.GetCanonicalMediaNameList(ps);
                if (mns.Contains(pageSize))
                { psv.SetCanonicalMediaName(ps, pageSize); }
                // Set the pen settings
                System.Collections.Specialized.StringCollection ssl = psv.GetPlotStyleSheetList();
                if (ssl.Contains(styleSheet))
                { psv.SetCurrentStyleSheet(ps, styleSheet); }
                // Copy the PlotSettings data back to the Layout
                bool upgraded = false;
                if (!lay.IsWriteEnabled)
                {
                    lay.UpgradeOpen();
                    upgraded = true;
                }
                lay.CopyFrom(ps);
                if (upgraded)
                { lay.DowngradeOpen(); }
            }

        }

        /// <summary>
        /// Determine the maximum possible size for this layout.
        /// </summary>
        /// <returns>The maximum extents of the viewport on this layout.</returns>
        public static Extents2d GetMaximumExtents(this Layout lay)
        {
            // If the drawing template is imperial, we need to divide by
            // 1" in mm (25.4)
            //double div = lay.PlotPaperUnits == PlotPaperUnit.Inches ? 25.4 : 1.0;
            double div = lay.PlotPaperUnits == PlotPaperUnit.Millimeters ? 1.0 : 1.0;
            // We need to flip the axes if the plot is rotated by 90 or 270 deg
            bool doIt = lay.PlotRotation == PlotRotation.Degrees090 || lay.PlotRotation == PlotRotation.Degrees270;
            // Get the extents in the correct units and orientation
            Point2d min = lay.PlotPaperMargins.MinPoint / div;
            Point2d max = (lay.PlotPaperSize - lay.PlotPaperMargins.MaxPoint.GetAsVector()) / div;
            return new Extents2d(min, max);
        }

        /// <summary>
        /// Sets the size of the viewport according to the provided extents.
        /// </summary>
        /// <param name="ext">The extents of the viewport on the page.</param>
        /// <param name="fac">Optional factor to provide padding.</param>
        public static void ResizeViewport(this Viewport vp, Extents2d ext, double fac = 1.0)
        {
            vp.Width = (ext.MaxPoint.X - ext.MinPoint.X) * fac;
            vp.Height = (ext.MaxPoint.Y - ext.MinPoint.Y) * fac;
            vp.CenterPoint = (Point2d.Origin + (ext.MaxPoint - ext.MinPoint) * 0.5).Pad();
        }

        /// <summary>
        /// Sets the view in a viewport to contain the specified model extents.
        /// </summary>
        /// <param name="ext">The extents of the content to fit the viewport.</param>
        /// <param name="fac">Optional factor to provide padding.</param>
        public static void FitContentToViewport(this Viewport vp, Extents3d ext, double fac = 1.0)
        {
            vp.CustomScale = fac;
            // Let's zoom to just larger than the extents
            vp.ViewCenter = (ext.MinPoint + ((ext.MaxPoint - ext.MinPoint) * 0.5)).Strip();


            // Get the dimensions of our view from the database extents
            double hgt = ext.MaxPoint.Y - ext.MinPoint.Y;
            double wid = ext.MaxPoint.X - ext.MinPoint.X;
            // We'll compare with the aspect ratio of the viewport itself
            // (which is derived from the page size)
            double aspect = vp.Width / vp.Height;
            // If our content is wider than the aspect ratio, make sure we
            // set the proposed height to be larger to accommodate the
            // content

            if (wid / hgt > aspect)
            { hgt = wid / aspect; }
            // Set the height so we're exactly at the extents
            vp.ViewHeight = hgt;
            // Set a custom scale to zoom out slightly (could also
            // vp.ViewHeight *= 1.1, for instance)

        }
    }

    public class LayoutByXrefBlock
    {
        [CommandMethod("LayByXrBl")]
        [Obsolete]
        public void LayoutbyXrefBlock()
        {
            Autodesk.AutoCAD.ApplicationServices.Document aDoc = Application.DocumentManager.MdiActiveDocument;
            Database db = aDoc.Database;
            Editor ed = aDoc.Editor;
            PromptStringOptions Prefix = new PromptStringOptions("Type Prefix: ");
            PromptResult prefixResult = ed.GetString(Prefix);
            PromptDoubleOptions number = new PromptDoubleOptions("Type number start: ");
            PromptDoubleResult startResult = ed.GetDouble(number);
            PromptSelectionResult psr = ed.GetSelection();
            SelectionSet ss = psr.Value;
            ObjectId objID = ss[0].ObjectId;
            using (Viewport oBlock = objID.Open(OpenMode.ForRead) as Viewport) 
            { }


            double v;
            if(startResult.Status!= PromptStatus.OK) { v = 0;}  
            else { v = startResult.Value; }
            //try
            //{
            //    //Chọn block mẫu
            //    string blockName;
            //    Point2d insertPoint;
            //    double oBScale;
            //    TypedValue[] tvs = new TypedValue[]
            //    {new TypedValue((int)DxfCode.Start, "INSERT") };
            //    SelectionFilter filter = new SelectionFilter(tvs);
            //    PromptSelectionResult psr = ed.GetSelection(filter);
            //    SelectionSet ss = psr.Value;
            //    ObjectId  objID  = ss[0].ObjectId;
            //    using(BlockReference oBlock = objID.Open(OpenMode.ForRead) as BlockReference)
            //    { 
            //        blockName = oBlock.Name;
            //        oBScale = oBlock.ScaleFactors.X;
            //        insertPoint = oBlock.Position.Strip();
            //    }                       
            //    //Chọn biên dạng của vùng khung tên
            //    PromptPointOptions p1click = new PromptPointOptions("Select Point P1: ");
            //    PromptPointResult p1Result = ed.GetPoint(p1click);
            //    PromptPointOptions p2click = new PromptPointOptions("Select Point P2: ");
            //    PromptPointResult p2Result = ed.GetPoint(p2click);
            //    Point2d p1 = p1Result.Value.Strip();
            //    Point2d p2 = p2Result.Value.Strip();

            //    //Set biên dạng cho các block 
            //    Vector2d InP1vector = p1 - insertPoint;
            //    Vector2d InP2vector = p2 - insertPoint;

            //    //Chọn tất cả các block dựa vào block mẫu
            //    TypedValue[] tvbyName = new TypedValue[]
            //    {new TypedValue((int)DxfCode.BlockName, blockName) };
            //    SelectionFilter filterbyName = new SelectionFilter(tvbyName);
            //    PromptSelectionResult psrbyName = ed.GetSelection(filterbyName);
            //    SelectionSet ssbyName = psrbyName.Value;
            //    List<BlockReference> listBlock = new List<BlockReference>();
            //    for (int i = 0; i < ssbyName.Count; i++)
            //    {
                    
            //        double dwgNo = startResult.Value + i;
            //        string name = null;
            //        if (dwgNo < 10)
            //        { name = "0" + dwgNo.ToString(); }
            //        else { name = dwgNo.ToString(); }
            //        string fulldwgname = prefixResult.StringResult + name;
            //        Extents2d ext = new Extents2d();
            //        using (Transaction tr = db.TransactionManager.StartTransaction())
            //        {
            //            ObjectId objIDbyName = ssbyName[i].ObjectId;
            //            using (BlockReference oblByName = objIDbyName.Open(OpenMode.ForRead) as BlockReference)
            //            {
            //                double objScale = oblByName.ScaleFactors.X;
            //                double scaleFactor = objScale / oBScale;
            //                Point2d objPosition = oblByName.Position.Strip();
            //                Vector2d P1scaleVecto = InP1vector * scaleFactor;
            //                Vector2d P2scaleVecto = InP2vector * scaleFactor;

            //                Point2d PointP1block = objPosition + P1scaleVecto;
            //                Point2d PointP2block = objPosition + P2scaleVecto;

            //                // Create and select a new layout tab
            //                ObjectId id = LayoutManager.Current.CreateAndMakeLayoutCurrent(fulldwgname);
            //                // Open the created layout
            //                Layout lay = (Layout)tr.GetObject(id, OpenMode.ForWrite);
            //                // Make some settings on the layout and get its extents
            //                _ = oblByName.Bounds.Value.MinPoint;
            //                lay.SetPlotSettings
            //                    (
            //                        //"ISO_full_bleed_2A0_(1189.00_x_1682.00_MM)", // Try this big boy!
            //                        "ISO_full_bleed_A1_(841.00_x_594.00_MM)",
            //                        "monochrome.ctb",
            //                        "DWF6 ePlot.pc3"
            //                    );
            //                ext = lay.GetMaximumExtents();
            //                lay.ApplyToViewport
            //                (
            //                    tr, 2,
            //                    vp =>
            //                    {
            //                        // Size the viewport according to the extents calculated when
            //                        // we set the PlotSettings (device, page size, etc.)
            //                        // Use the standard 10% margin around the viewport
            //                        // (found by measuring pixels on screenshots of Layout1, etc.)
            //                        vp.ResizeViewport(ext, 1);
            //                        // Adjust the view so that the model contents fit
            //                        //if (ValidDbExtents(minPoint3D, maxPoint3D))
            //                        vp.FitContentToViewport(new Extents3d(PointP1block.Pad(), PointP2block.Pad()), 1 / (oblByName.ScaleFactors.X));
            //                        // Finally we lock the view to prevent meddling
            //                        vp.Locked = true;
            //                    }
            //                );
            //                //Commit the transaction
            //                tr.Commit();
            //            }
            //            // Zoom so that we can see our new layout, again with a little padding
            //            ed.Command("_.ZOOM", "_E");
            //            ed.Regen();
            //        }
                    
            //    }    
                

            //}
            //catch (Exception) 
            //{ 
            //}
        }
        // Returns whether the provided DB extents - retrieved from
        // Database.Extmin/max - are "valid" or whether they are the default
        // invalid values (where the min's coordinates are positive and the
        // max coordinates are negative)
        private bool ValidDbExtents(Point3d min, Point3d max)
        {
            return
              !(min.X > 0 && min.Y > 0 && min.Z > 0 &&
                max.X < 0 && max.Y < 0 && max.Z < 0);

        }
    }
}
