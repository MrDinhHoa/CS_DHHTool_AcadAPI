using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace _00_TestCommand
{
    public class TestCommand
    {
        Document doc;  // active document
        double radius; // radius default value
        string layer;  // layer default value

        /// <summary>
        /// Creates a new instance of Commands.
        /// This constructor is run once per document
        /// at the first call of a 'CommandMethod' method.
        /// </summary>
        public TestCommand()
        {
            // private fields initialization (initial default values)
            doc = AcAp.DocumentManager.MdiActiveDocument;
            radius = 10.0;
            layer = (string)AcAp.GetSystemVariable("clayer");
        }
        [CommandMethod("CMD_CIRCLE")]
        public void DrawCircleCmd()
        {
            var db = doc.Database;
            var ed = doc.Editor;

            // choose of the layer
            var layers = GetLayerNames(db);
            if (!layers.Contains(layer))
            {
                layer = (string)AcAp.GetSystemVariable("clayer");
            }
            var strOptions = new PromptStringOptions("\nLayer name: ");
            strOptions.DefaultValue = layer;
            strOptions.UseDefaultValue = true;
            var strResult = ed.GetString(strOptions);
            if (strResult.Status != PromptStatus.OK)
                return;
            if (!layers.Contains(strResult.StringResult))
            {
                ed.WriteMessage(
                  $"\nNone layer named '{strResult.StringResult}' in the layer table.");
                return;
            }
            layer = strResult.StringResult;

            // specification of the radius
            var distOptions = new PromptDistanceOptions("\nSpecify the radius: ");
            distOptions.DefaultValue = radius;
            distOptions.UseDefaultValue = true;
            var distResult = ed.GetDistance(distOptions);
            if (distResult.Status != PromptStatus.OK)
                return;
            radius = distResult.Value;

            // specification of the center
            var ppr = ed.GetPoint("\nSpecify the center: ");
            if (ppr.Status == PromptStatus.OK)
            {
                // drawing of the circle in the current space
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var curSpace =
                      (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    using (var circle = new Circle(ppr.Value, Vector3d.ZAxis, radius))
                    {
                        circle.TransformBy(ed.CurrentUserCoordinateSystem);
                        circle.Layer = strResult.StringResult;
                        curSpace.AppendEntity(circle);
                        tr.AddNewlyCreatedDBObject(circle, true);
                    }
                    tr.Commit();
                }
            }
        }

        /// <summary>
        /// Gets the layer list.
        /// </summary>
        /// <param name="db">Database instance this method applies to.</param>
        /// <returns>Layer names list.</returns>
        private List<string> GetLayerNames(Database db)
        {
            var layers = new List<string>();
            using (var tr = db.TransactionManager.StartOpenCloseTransaction())
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId id in layerTable)
                {
                    var layer = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    layers.Add(layer.Name);
                }
            }
            return layers;
        }
        [CommandMethod("MyPick")]
        public static void RunMyCommand()
        {
            Document dwg = Application.DocumentManager.MdiActiveDocument;
            Editor ed = dwg.Editor;

            int val = 0;
            ObjectId pickedId = ObjectId.Null;

            while (true)
            {
                PromptEntityOptions opt = new PromptEntityOptions(
                    "\nSelect an entity or [Set value] <30>:");

                opt.AllowNone = true;
                opt.Keywords.Add("Set value");
                opt.Keywords.Add("30");
                opt.Keywords.Default = "30";
                opt.AppendKeywordsToMessage = false;

                PromptEntityResult res = ed.GetEntity(opt);
                if (res.Status == PromptStatus.OK)
                {
                    ed.WriteMessage("\nPicked entity: {0}", res.ObjectId.ToString());
                    pickedId = res.ObjectId;
                    break;
                }
                else if (res.Status == PromptStatus.Keyword)
                {
                    ed.WriteMessage("keyword: {0}", res.StringResult);
                    if (res.StringResult == "Set")
                    {
                        //use Editor.GetInteger() tp get new value
                        //val=....
                    }
                    else
                    {
                        val = 30;
                    }
                }
                else
                {
                    ed.WriteMessage("\nInvalid pick or cancelled");
                    break;
                }
            }

            if (pickedId != ObjectId.Null)
            {
                //Do something
            }

            Autodesk.AutoCAD.Internal.Utils.PostCommandPrompt();
        }

        [CommandMethod("MyCmd", CommandFlags.Session)]
        public static void RunMy2Command()
        {
            Document dwg = Application.DocumentManager.MdiActiveDocument;
            Editor ed = dwg.Editor;
            PromptEntityOptions opt = new PromptEntityOptions("\nSelect an entity:");
            opt.AllowNone = true;
            opt.Keywords.Add("Aaaa");
            opt.Keywords.Add("Bbbb");
            opt.Keywords.Add("Cccc");
            opt.Keywords.Default = "Cccc";
            opt.Keywords[2].Visible = false;

            PromptEntityResult res = ed.GetEntity(opt);
            if (res.Status == PromptStatus.OK)
            {
                ed.WriteMessage("\nSelected entity: {0}", res.ObjectId.ToString());
            }
            else if (res.Status == PromptStatus.Keyword)
            {
                ed.WriteMessage("Keyword entered: {0}", res.StringResult);
            }
            else
            {
                ed.WriteMessage("\n*Cancell*");
            }

            ed.WriteMessage("\nMyCmd executed");
            Autodesk.AutoCAD.Internal.Utils.PostCommandPrompt();
        }
    }

}
