using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace _03_DrawSectionBeam
{
    class Library
    {
        public static Polyline drawRectangle (Point2d insertPoint, double width, double height)
        {
            Polyline rectangle = new Polyline();
            rectangle.AddVertexAt(0, new Point2d(insertPoint.X, insertPoint.Y), 0, 0, 0);
            rectangle.AddVertexAt(1, new Point2d(insertPoint.X + width, insertPoint.Y), 0, 0, 0);
            rectangle.AddVertexAt(2, new Point2d(insertPoint.X + width, insertPoint.Y + height), 0, 0, 0);
            rectangle.AddVertexAt(3, new Point2d(insertPoint.X, insertPoint.Y + height), 0, 0, 0);
            rectangle.AddVertexAt(4, new Point2d(insertPoint.X, insertPoint.Y), 0, 0, 0);
            return rectangle;
        }
        public static Polyline drawstirrup(Point2d insertPoint, double width, double height, double radius, double offset)
        {
            Polyline stirrup = new Polyline();
            double bugle = Math.Tan(Math.PI / 8);
            stirrup.AddVertexAt(0, new Point2d(insertPoint.X + offset + radius, insertPoint.Y + offset), 0, 0, 0);
            stirrup.AddVertexAt(0, new Point2d(insertPoint.X + width - offset - radius, insertPoint.Y + offset), 0, 0, 0);
            stirrup.AddVertexAt(0, new Point2d(insertPoint.X + width - offset, insertPoint.Y + radius + offset), -bugle, 0, 0);
            stirrup.AddVertexAt(0, new Point2d(insertPoint.X + width - offset, insertPoint.Y + height - radius - offset), 0, 0, 0);
            stirrup.AddVertexAt(0, new Point2d(insertPoint.X + width - radius - offset, insertPoint.Y + height - offset), -bugle, 0, 0);
            stirrup.AddVertexAt(0, new Point2d(insertPoint.X + radius + offset, insertPoint.Y + height - offset), 0, 0, 0);
            stirrup.AddVertexAt(0, new Point2d(insertPoint.X + offset, insertPoint.Y + height - radius - offset), -bugle, 0, 0);
            stirrup.AddVertexAt(0, new Point2d(insertPoint.X + offset, insertPoint.Y + radius + offset), 0, 0, 0);
            stirrup.AddVertexAt(0, new Point2d(insertPoint.X + radius + offset, insertPoint.Y + offset), -bugle, 0, 0);
            return stirrup;
        }

        public static Polyline drawhookRebar(Point2d insertPointHookRebar, double width, double offset)
        {
            Polyline HookRebar = new Polyline();
            double bugle = Math.Tan(Math.PI / 8);
            HookRebar.AddVertexAt(0,new Point2d(insertPointHookRebar.X, insertPointHookRebar.Y), 0, 0,0);
            HookRebar.AddVertexAt(0, new Point2d(insertPointHookRebar.X - 23, insertPointHookRebar.Y), 0, 0, 0);
            HookRebar.AddVertexAt(0, new Point2d(insertPointHookRebar.X - 23 - 17, insertPointHookRebar.Y - 17.5), -bugle, 0, 0);
            HookRebar.AddVertexAt(0, new Point2d(insertPointHookRebar.X -23, insertPointHookRebar.Y - 35), -bugle, 0, 0);
            HookRebar.AddVertexAt(0, new Point2d(insertPointHookRebar.X - 23 + width - 2 * offset - 34, insertPointHookRebar.Y - 35),0, 0, 0);
            HookRebar.AddVertexAt(0, new Point2d(insertPointHookRebar.X - 23 + width - 2 * offset - 34 + 17, insertPointHookRebar.Y - 17.5), -bugle, 0, 0);
            HookRebar.AddVertexAt(0, new Point2d(insertPointHookRebar.X - 23 + width - 2 * offset - 34, insertPointHookRebar.Y), -bugle, 0, 0);
            HookRebar.AddVertexAt(0, new Point2d(insertPointHookRebar.X - 23 + width - 2 * offset - 34 - 23, insertPointHookRebar.Y),0, 0, 0);
            return HookRebar;
        }
    }
}
