using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using netDxf;
using System.Collections.ObjectModel;
using netDxf.Entities;
using System.Windows.Controls;
using netDxf.Blocks;
using System.Windows.Media;

namespace DynamicGeometry
{
    public class DXFDrawingDeserializer
    {

        Drawing drawing;
        DxfDocument doc;

        public Drawing ReadDrawing(string dxfFileName, Canvas canvas)
        {
            doc = new DxfDocument();
            doc.Load(dxfFileName);

            drawing = new Drawing(canvas);

            ReadLines();
            ReadPolylines();
            ReadArcs();
            ReadCircles();
            ReadInserts();

            drawing.Recalculate();
            return drawing;
        }

        FreePoint CreateHiddenPoint(double x, double y)
        {
            return new FreePoint()
            {
                Drawing = drawing,
                X = x,
                Y = y,
                Visible = false
            };
        }

        Segment CreateSegment(IPoint p1, IPoint p2)
        {
            return Factory.CreateSegment(drawing, new[] { p1, p2 });
        }

        netDxf.Entities.Polyline CastPolyline(IPolyline item)
        {
            netDxf.Entities.Polyline polyline = null;
            if (item is LightWeightPolyline)
            {
                polyline = ((LightWeightPolyline)item).ToPolyline();
            }
            else if (item is Polyline)
            {
                polyline = (netDxf.Entities.Polyline)item;
            }
            else
            {
                polyline = null;
            }
            return polyline;
        }

        void ReadLines()
        {
            foreach (var line in doc.Lines)
            {
                ReadLine(line, 0, 0);
            }
        }

        void ReadLine(Line line, double x, double y)
        {
            var point1 = CreateHiddenPoint(line.StartPoint.X + x, line.StartPoint.Y + y);
            var point2 = CreateHiddenPoint(line.EndPoint.X + x, line.EndPoint.Y + y);
            var segment = CreateSegment(point1, point2);
            Actions.Add(drawing, segment);
        }

        void ReadPolylines()
        {
            foreach (var item in doc.Polylines)
            {
                netDxf.Entities.Polyline polyline = CastPolyline(item);
                if (polyline != null)
                {
                    ReadPolyline(polyline.Vertexes, polyline.IsClosed, 0, 0);
                }
            }
        }

        void ReadPolyline(IList<PolylineVertex> vertices, bool isClosed, double x, double y)
        {
            IPoint firstPoint = null;
            IPoint previousPoint = null;
            var figures = new List<IFigure>();
            var segments = new List<IFigure>();

            foreach (var vertex in vertices)
            {
                var point = CreateHiddenPoint(vertex.Location.X + x, vertex.Location.Y + y);
                if (firstPoint == null)
                {
                    firstPoint = point;
                }
                if (previousPoint != null)
                {
                    var segment = CreateSegment(previousPoint, point);
                    figures.Add(segment);
                    segments.Add(segment);
                }
                previousPoint = point;
                figures.Add(point);
            }
            if (previousPoint != null && isClosed)
            {
                var segment = CreateSegment(previousPoint, firstPoint);
                figures.Add(segment);
                segments.Add(segment);

                var polygon = Factory.CreatePolygon(drawing, figures);
                Actions.Add(drawing, polygon);
            }

            Actions.AddMany(drawing, segments.ToArray());
        }

        void ReadArcs()
        {
            foreach (var item in doc.Arcs)
            {
                ReadArc(item, 0, 0);
            }
        }

        void ReadArc(netDxf.Entities.Arc arc, double x, double y)
        {
            // TODO :  
        }

        void ReadCircles()
        {
            foreach (var item in doc.Circles)
            {
                ReadCircle(item, 0, 0);
            }
        }

        void ReadCircle(netDxf.Entities.Circle circle, double x, double y)
        {
            var figures = new List<IFigure>();

            figures.Add(CreateHiddenPoint(circle.Center.X + x, circle.Center.Y + y));
            figures.Add(CreateHiddenPoint(circle.Center.X + x + circle.Radius, circle.Center.Y + y));

            var figure = Factory.CreateCircleByRadius(drawing, figures);
        }

        void ReadInserts()
        {
            foreach (var item in doc.Inserts)
            {
                ReadInsert(item);
            }
        }

        void ReadInsert(netDxf.Entities.Insert insert)
        {
            List<netDxf.Entities.IEntityObject> entities = insert.Block.Entities;
            netDxf.Entities.IEntityObject entity = null;

            for (int index = 1; index < entities.Count; index++)
            {
                entity = entities[index];

                if (entity is Line)
                    ReadLine((Line)entity, insert.InsertionPoint.X, insert.InsertionPoint.Y);
                else if (entity is netDxf.Entities.Arc)
                    ReadArc((netDxf.Entities.Arc)entity, insert.InsertionPoint.X, insert.InsertionPoint.Y);
                else if (entity is netDxf.Entities.Circle)
                    ReadCircle((netDxf.Entities.Circle)entity, insert.InsertionPoint.X, insert.InsertionPoint.Y);
                else if (entity is IPolyline)
                {
                    netDxf.Entities.Polyline polyline = CastPolyline((IPolyline)entity);
                    if (polyline != null)
                    {
                        ReadPolyline(polyline.Vertexes, polyline.IsClosed, 0, 0);
                    }
                }

            }

        }

    }
}
