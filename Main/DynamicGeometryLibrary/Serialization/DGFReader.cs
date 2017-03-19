using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;

namespace DynamicGeometry
{
    public class DGFReader
    {
        public bool IsSuccess
        {
            get
            {
                return true;
            }
        }

        public string GetErrorReport()
        {
            return "";
        }

        public void ReadDrawing(Drawing drawing, string[] lines)
        {
            this.drawing = drawing;

            var ini = new IniFile(lines);

            foreach (var section in ini.Sections)
            {
                sectionLookup.Add(section.Title, section);
            }

            using (drawing.ActionManager.CreateTransaction())
            {
                drawing.ActionManager.RecordingTransaction.IsDelayed = false;
                foreach (var section in ini.Sections)
                {
                    ProcessSection(section);
                }
            }
            DrawingUpdater updater;
#if TABULA
            updater = new TABDrawingUpdater();
#else
            updater = new DrawingUpdater();
#endif
            updater.UpdateIfNecessary(drawing);
            drawing.Recalculate();
        }

        void EnsureSectionProcessed(string sectionTitle)
        {
            if (processedSections.ContainsKey(sectionTitle))
            {
                return;
            }

            var section = sectionLookup[sectionTitle];
            ProcessSection(section);
        }

        Dictionary<string, IniFile.Section> sectionLookup = new Dictionary<string, IniFile.Section>();
        Dictionary<string, IniFile.Section> processedSections = new Dictionary<string, IniFile.Section>();

        Drawing drawing;

        void ProcessSection(IniFile.Section section)
        {
            string title = section.Title;
            if (processedSections.ContainsKey(title))
            {
                return;
            }

            processedSections.Add(title, section);
            if (title == "General")
            {
                ProcessGeneral(section);
            }
            else if (title.StartsWith("Point"))
            {
                ProcessPoint(section);
            }
            else if (title.StartsWith("Figure"))
            {
                ProcessFigure(section);
            }
            else if (title.StartsWith("Label"))
            {
                ProcessLabel(section);
            }
            else if (title.StartsWith("Locus"))
            {
                ProcessLocus(section);
            }
            else if (title.StartsWith("SG"))
            {
                ProcessSG(section);
            }
            else if (title.StartsWith("Button"))
            {
                ProcessButton(section);
            }
        }

        void ProcessButton(IniFile.Section section)
        {
            int type = section.ReadInt("Type");
            if (type != 0)
            {
                return;
            }

            ShowHideControl button = new ShowHideControl();
            button.Drawing = drawing;
            button.Checkbox.IsChecked = section.ReadBool("InitiallyVisible", false);
            button.AddDependencies(GetPointList(section));
            button.MoveTo(section.ReadDouble("X"), section.ReadDouble("Y"));
            SetButtonStyle(section, button);
            Actions.Add(drawing, button);
            
            string text = section["Caption"].Replace("~~", Environment.NewLine);
            if (section["Charset"] == "0")
            {
                text = text.Replace('I', '²');
            }

            button.Checkbox.Content = text;

            ReadButtonDependencies(section, button);

            button.UpdateFigureVisibility();
            buttons[section.GetTitleNumber("Button")] = button;
        }

        private void ReadButtonDependencies(IniFile.Section section, ShowHideControl button)
        {
            List<IFigure> dependencies = new List<IFigure>();

            AddButtonDependencies(section, dependencies, "Button", buttons);
            AddButtonDependencies(section, dependencies, "Figure", figures);
            AddButtonDependencies(section, dependencies, "Label", labels);
            AddButtonDependencies(section, dependencies, "Point", points);
            AddButtonDependencies(section, dependencies, "SG", staticGraphics);

            button.AddDependencies(dependencies);
        }

        private void AddButtonDependencies(
            IniFile.Section section, 
            List<IFigure> dependencies,
            string dependencyType,
            IList container)
        {
            foreach (int dependency in ReadIndices(section, dependencyType))
            {
                EnsureSectionProcessed(dependencyType + dependency.ToString());
                dependencies.Add((IFigure)container[dependency]);
            }
        }

        private IEnumerable<int> ReadIndices(IniFile.Section section, string key)
        {
            int index = 1;

            string currentKey = key + "#" + index.ToString();
            while (section.Entries.ContainsKey(currentKey))
            {
                yield return section.ReadInt(currentKey);
                index++;
                currentKey = key + "#" + index.ToString();
            }
        }

        void ProcessSG(IniFile.Section section)
        {
            var type = section.ReadInt("Type");

            switch (type)
            {
                case 0: // Polygon
                    ReadPolygon(section);
                    break;
                case 1: // Bezier
                    ReadBezier(section);
                    break;
                case 2: // Vector
                    ReadVector(section);
                    break;
                default:
                    break;
            }
        }

        void ReadVector(IniFile.Section section)
        {
            var vector = Factory.CreateVector(
                drawing, GetPointList(section));
            Actions.Add(drawing, vector);
            SetFigureStyle(section, vector);
            staticGraphics[section.GetTitleNumber("SG")] = vector;
        }

        void ReadBezier(IniFile.Section section)
        {
            var bezier = Factory.CreateBezier(
                drawing, GetPointList(section));
            Actions.Add(drawing, bezier);
            SetFigureStyle(section, bezier);
            staticGraphics[section.GetTitleNumber("SG")] = bezier;
        }

        void ReadPolygon(IniFile.Section section)
        {
            var polygon = Factory.CreatePolygon(
                drawing, GetPointList(section));
            Actions.Add(drawing, polygon);
            SetFigureStyle(section, polygon);
            staticGraphics[section.GetTitleNumber("SG")] = polygon;
        }

        void ProcessLocus(IniFile.Section section)
        {
        }

        void ProcessLabel(IniFile.Section section)
        {
            var label = Factory.CreateLabel(drawing);
            label.MoveTo(section.ReadDouble("X"), section.ReadDouble("Y"));
            SetLabelStyle(section, label);
            Actions.Add(drawing, label);
            labels[section.GetTitleNumber("Label")] = label;
            string text = section["Caption"].Replace("~~", Environment.NewLine);
            if (section["Charset"] == "0")
            {
                text = text.Replace('І', '²');
            }
#if !PLAYER
            drawing.ActionManager.SetProperty(label, "Text", text);
#else
            label.Text = text;
#endif
        }

        void ProcessFigure(IniFile.Section section)
        {
            var figureType = section.ReadInt("FigureType");

            switch (figureType)
            {
                case 1: // Segment
                    ReadSegment(section);
                    break;
                case 2: // Ray
                    ReadRay(section);
                    break;
                case 3: // Line
                    ReadLine2Points(section);
                    break;
                case 4: // ParallelLine
                    ReadParallelLine(section);
                    break;
                case 5: // PerpendicularLine
                    ReadPerpendicularLine(section);
                    break;
                case 6: // Circle
                    ReadCircle(section);
                    break;
                case 7: // CircleByRadius
                    ReadCircleByRadius(section);
                    break;
                case 8: // Arc
                    ReadArc(section);
                    break;
                case 9: // MidPoint
                    ReadMidPoint(section);
                    break;
                case 10: // SymmPoint
                    ReadSymmPoint(section);
                    break;
                case 11: // SymmPointByLine
                    ReadSymmPointByLine(section);
                    break;
                case 12: // InvertedPoint
                    ReadInvertedPoint(section);
                    break;
                case 13: // IntersectionPoint
                    ReadIntersectionPoint(section);
                    break;
                case 14: // PointOnFigure
                    ReadPointOnFigure(section);
                    break;
                case 15: // MeasureDistance
                    ReadMeasureDistance(section);
                    break;
                case 16: // MeasureAngle
                    ReadMeasureAngle(section);
                    break;
                case 17: // AnLineGeneral
                    ReadAnalyticLineGeneral(section);
                    break;
                case 21: // AnCircle
                    ReadAnalyticCircle(section);
                    break;
                case 22: // AnalyticPoint
                    ReadAnalyticPoint(section);
                    break;
                case 23: // Locus
                    ReadLocus(section);
                    break;
                case 24: // AngleBisector
                    ReadAngleBisector(section);
                    break;
                default:
                    break;
            }
        }

        void ReadAnalyticCircle(IniFile.Section section)
        {
            double a = section.ReadDouble("AuxInfo(1)");
            double b = section.ReadDouble("AuxInfo(2)");
            double c = section.ReadDouble("AuxInfo(3)");

            double x = -a / 2;
            double y = -b / 2;
            double r = (a * a / 4 + b * b / 4 - c).SquareRoot();

            var figure = Factory.CreateCircleByEquation(
                drawing,
                x.ToStringInvariant(),
                y.ToStringInvariant(),
                r.ToStringInvariant());
            SetFigureStyle(section, figure);
            AddFigure(section, figure);
        }

        void ReadAnalyticLineGeneral(IniFile.Section section)
        {
            var figure = Factory.CreateLineByEquation(
                drawing,
                section["AuxInfo(1)"],
                section["AuxInfo(2)"],
                section["AuxInfo(3)"]);
            SetFigureStyle(section, figure);
            AddFigure(section, figure);
        }

        void ReadAnalyticPoint(IniFile.Section section)
        {
            int numberOfPoints = section.ReadInt("NumberOfPoints");
            for (int i = 1; i < numberOfPoints; i++)
            {
                GetPoint(section, i);
            }

            string x = section["XS"];
            string y = section["YS"];

            var point = Factory.CreatePointByCoordinates(drawing, x, y);
            AddFigure(section, point);
            AddFigurePoint(section, point, 0);
        }

        void ReadMeasureAngle(IniFile.Section section)
        {
            var vertex = GetPoint(section, 1);
            var point1 = GetPoint(section, 0);
            var point2 = GetPoint(section, 2);

            if (Math.OAngle(point1.Coordinates, vertex.Coordinates, point2.Coordinates) > Math.PI)
            {
                var temp = point1;
                point1 = point2;
                point2 = temp;
            }
            var angle = Factory.CreateAngleMeasurement(
                drawing,
                new[] {
                    vertex,
                    point1,
                    point2});
            if (section["AuxPoint(6)"] != null)
            {
                angle.Offset = drawing.CoordinateSystem.ToLogical(GetAuxPoint(section, 6));
            }

            AddLabelFigure(section, angle);
        }

        void ReadIntersectionPoint(IniFile.Section section)
        {
            var pointSection1 = sectionLookup["Point" + section["Points0"]];
            var pointSection2 = sectionLookup["Point" + section["Points1"]];

            var point1 = Factory.CreateIntersectionPoint(
                drawing,
                GetParent(section, 0),
                GetParent(section, 1),
                new Point(pointSection1.ReadDouble("X"), pointSection1.ReadDouble("Y")));

            var point2 = Factory.CreateIntersectionPoint(
                drawing,
                GetParent(section, 0),
                GetParent(section, 1),
                new Point(pointSection2.ReadDouble("X"), pointSection2.ReadDouble("Y")));

            AddFigure(section, point1);
            Actions.Add(drawing, point2);
            AddFigurePoint(section, point1, 0);
            AddFigurePoint(section, point2, 1);
        }

        void ReadInvertedPoint(IniFile.Section section)
        {
            var point = Factory.CreateReflectedPoint(
                drawing, new[] { GetPoint(section, 1), GetParent(section, 0) });
            AddFigure(section, point);
            AddFigurePoint(section, point, 0);
        }

        void ReadSymmPointByLine(IniFile.Section section)
        {
            var point = Factory.CreateReflectedPoint(
                drawing, new[] { GetPoint(section, 1), GetParent(section, 0) });
            AddFigure(section, point);
            AddFigurePoint(section, point, 0);
        }

        void ReadArc(IniFile.Section section)
        {
            var radiusPoint1 = GetPoint(section, 0);
            var radiusPoint2 = GetPoint(section, 1);
            var centerPoint = GetPoint(section, 2);
            var anglePoint1 = GetPoint(section, 3);
            var anglePoint2 = GetPoint(section, 4);

            var endPoint1 = Factory.CreatePointByCoordinates(
                drawing,
                "{2}.X + ({3}.X - {2}.X) * {0}{1} / {2}{3}"
                .Format(radiusPoint1, radiusPoint2, centerPoint, anglePoint1),
                "{2}.Y + ({3}.Y - {2}.Y) * {0}{1} / {2}{3}"
                .Format(radiusPoint1, radiusPoint2, centerPoint, anglePoint1));
            var endPoint2 = Factory.CreatePointByCoordinates(
                drawing,
                "{2}.X + ({3}.X - {2}.X) * {0}{1} / {2}{3}"
                .Format(radiusPoint1, radiusPoint2, centerPoint, anglePoint2),
                "{2}.Y + ({3}.Y - {2}.Y) * {0}{1} / {2}{3}"
                .Format(radiusPoint1, radiusPoint2, centerPoint, anglePoint2));

            Actions.Add(drawing, endPoint1);
            Actions.Add(drawing, endPoint2);

            AddFigurePoint(section, endPoint1, 5);
            AddFigurePoint(section, endPoint2, 6);

            var figure = Factory.CreateArc(
                drawing, new[] { centerPoint, endPoint1, endPoint2 });
            SetFigureStyle(section, figure);
            AddFigure(section, figure);
        }

        void ReadCircleByRadius(IniFile.Section section)
        {
            var figure = Factory.CreateCircleByRadius(
                drawing, GetPointList(section, 3));
            SetFigureStyle(section, figure);
            AddFigure(section, figure);
        }

        void ReadAngleBisector(IniFile.Section section)
        {
            var figure = Factory.CreateAngleBisector(
                drawing, new[] {
                    GetPoint(section, 1),
                    GetPoint(section, 0),
                    GetPoint(section, 2)});
            SetFigureStyle(section, figure);
            AddFigure(section, figure);
        }

        void ReadLocus(IniFile.Section section)
        {
            var figure = Factory.CreateLocus(
                drawing, GetPointList(section, 2));
            var locusSection = sectionLookup
                ["Locus" + sectionLookup["Point" + section["Points0"]]["Locus"]];
            SetFigureStyle(locusSection, figure);
            AddFigure(section, figure);
        }

        void ReadSymmPoint(IniFile.Section section)
        {
            var point = Factory.CreateReflectedPoint(
                drawing, new[] {
                    GetPoint(section, 1),
                    GetPoint(section, 2)});
            AddFigure(section, point);
            AddFigurePoint(section, point, 0);
        }

        void ReadMidPoint(IniFile.Section section)
        {
            var point = Factory.CreateMidPoint(
                drawing, new[] {
                    GetPoint(section, 1),
                    GetPoint(section, 2)});
            AddFigure(section, point);
            AddFigurePoint(section, point, 0);
        }

        void ReadPointOnFigure(IniFile.Section section)
        {
            var point = Factory.CreatePointOnFigure(
                drawing, GetParent(section, 0), new Point());
            point.Parameter = -GetAuxInfo(section, 1);
            AddFigure(section, point);
            AddFigurePoint(section, point, 0);
        }

        void AddFigurePoint(IniFile.Section section, PointBase point, int index)
        {
            var pointSection1 = sectionLookup["Point" + section["Points" + index.ToString()]];
            point.Name = pointSection1["Name"];
            point.X = pointSection1.ReadDouble("X");
            point.Y = pointSection1.ReadDouble("Y");
            SetPointStyle(pointSection1, point);
            var pointIndex = GetPointIndex(section, index);
            points[pointIndex] = point;
        }

        void ReadMeasureDistance(IniFile.Section section)
        {
            var figure = Factory.CreateDistanceMeasurement(
                drawing, GetPointList(section, 2));
            if (section["AuxPoint(6)"] != null)
            {
                figure.Offset = drawing.CoordinateSystem.ToLogical(GetAuxPoint(section, 6));
            }
            AddLabelFigure(section, figure);
        }

        void ReadCircle(IniFile.Section section)
        {
            var figure = Factory.CreateCircle(
                drawing, GetPointList(section, 2));
            SetFigureStyle(section, figure);
            AddFigure(section, figure);
        }

        void ReadPerpendicularLine(IniFile.Section section)
        {
            var figure = Factory.CreatePerpendicularLine(
                drawing, new[] {
                    GetParent(section, 0),
                    GetPoint(section, 0)});
            SetFigureStyle(section, figure);
            AddFigure(section, figure);
        }

        void ReadParallelLine(IniFile.Section section)
        {
            var figure = Factory.CreateParallelLine(
                drawing, new[] {
                    GetParent(section, 0),
                    GetPoint(section, 0)});
            SetFigureStyle(section, figure);
            AddFigure(section, figure);
        }

        void ReadLine2Points(IniFile.Section section)
        {
            var figure = Factory.CreateLineTwoPoints(
                drawing, GetPointList(section, 2));
            SetFigureStyle(section, figure);
            AddFigure(section, figure);
        }

        IList<IFigure> GetPointList(IniFile.Section section, int numberOfPoints)
        {
            var result = new IFigure[numberOfPoints];
            for (int i = 0; i < numberOfPoints; i++)
            {
                result[i] = GetPoint(section, i);
            }

            var figureList = new List<IFigure>(result);
            return figureList;
        }

        IList<IFigure> GetPointList(IniFile.Section section)
        {
            int numberOfPoints = section.ReadInt("NumberOfPoints");
            var result = new List<IFigure>();
            for (int i = 1; i <= numberOfPoints; i++)
            {
                var point = GetPoint(section, i);
                result.Add(point);
            }

            return result;
        }

        void ReadRay(IniFile.Section section)
        {
            var figure = Factory.CreateRay(
                drawing, GetPointList(section, 2));
            SetFigureStyle(section, figure);
            AddFigure(section, figure);
        }

        void ReadSegment(IniFile.Section section)
        {
            var figure = Factory.CreateSegment(
                drawing, GetPointList(section, 2));
            SetFigureStyle(section, figure);
            AddFigure(section, figure);
        }

        void AddFigure(IniFile.Section section, IFigure figure)
        {
            Actions.Add(drawing, figure);
            figures[section.GetTitleNumber("Figure")] = figure;
        }

        void AddLabelFigure(IniFile.Section section, IFigure figure)
        {
            //SetLabelStyle(section, figure);
            Actions.Add(drawing, figure);
            figures[section.GetTitleNumber("Figure")] = figure;
        }

        void SetFigureStyle(IniFile.Section section, IFigure figure)
        {
            figure.Visible = !section.ReadBool("Hide", false) && section.ReadBool("Visible", true);

            var fillColor = section.ReadColor("FillColor");
            var foreColor = section.ReadColor("ForeColor");
            double drawWidth = section.TryReadDouble("DrawWidth") ?? 1.0;
            if (drawWidth < 0.1)
            {
                drawWidth = 0.1;
            }
            int? drawStyle = section.TryReadInt("DrawStyle");
            int? fillStyle = section.TryReadInt("FillStyle");
            int? drawMode = section.TryReadInt("DrawMode");
            if (fillStyle == null || fillStyle == 6)
            {
                fillColor = Colors.Transparent;
            }
            DoubleCollection strokeDashArray = null;
            if (drawStyle != null && drawStyle != 0)
            {
                strokeDashArray = new DoubleCollection();
                switch (drawStyle)
                {
                    case 1:
                        strokeDashArray.Add(15 / drawWidth, 6 / drawWidth);
                        break;
                    case 2:
                        strokeDashArray.Add(3 / drawWidth, 3 / drawWidth);
                        break;
                    case 3:
                        strokeDashArray.Add(10 / drawWidth, 4 / drawWidth, 2 / drawWidth, 4 / drawWidth);
                        break;
                    case 4:
                        strokeDashArray.Add(10 / drawWidth, 4 / drawWidth, 2 / drawWidth, 4 / drawWidth, 2 / drawWidth, 4 / drawWidth);
                        break;
                    default:
                        break;
                }
            }

            IFigureStyle style;

            if (figure is IShapeWithInterior)
            {
                if (fillColor != Colors.Transparent && drawMode != 13)
                {
                    fillColor.A = 128;
                }
                style = new ShapeStyle()
                {
                    Fill = new SolidColorBrush(fillColor),
                    Color = foreColor,
                    StrokeWidth = drawWidth,
                    StrokeDashArray = strokeDashArray
                };
            }
            else
            {
                if (foreColor == Colors.Transparent)
                {
                    System.Diagnostics.Debugger.Break();
                }
                style = new LineStyle()
                {
                    Color = foreColor,
                    StrokeWidth = drawWidth,
                    StrokeDashArray = strokeDashArray
                };
            }

            figure.Style = drawing.StyleManager.FindExistingOrAddNew(style);
        }

        void SetPointStyle(IniFile.Section section, IFigure figure)
        {
            figure.Visible = !section.ReadBool("Hide", false) && section.ReadBool("Visible", true);

            var fillColor = section.ReadColor("FillColor");
            var fillStyle = section.ReadInt("FillStyle");
            var foreColor = section.ReadColor("ForeColor");
            var physicalWidth = section.ReadDouble("PhysicalWidth");
            if (fillStyle == 1)
            {
                fillColor = Colors.Transparent;
            }

            PointStyle pointStyle = new PointStyle()
            {
                Fill = new SolidColorBrush(fillColor),
                Color = foreColor,
                Size = physicalWidth,
                StrokeWidth = 0.7
            };

            figure.Style = drawing.StyleManager.FindExistingOrAddNew(pointStyle);
        }

        void SetLabelStyle(IniFile.Section section, IFigure figure)
        {
            figure.Visible = !section.ReadBool("Hide", false) && section.ReadBool("Visible", true);

            var fontName = section["FontName"];
            double? fontSize = section.TryReadDouble("FontSize");
            bool fontUnderline = section.ReadBool("FontUnderline", false);
            var foreColor = section.ReadColor("ForeColor");
            var backColor = section.ReadColor("BackColor");

            TextStyle style = new TextStyle()
            {
                FontFamily = new FontFamily(fontName),
                Color = foreColor,
                Underline = fontUnderline
            };
            if (fontSize.HasValue)
            {
                style.FontSize = fontSize.Value * 1.7;
            }

            figure.Style = drawing.StyleManager.FindExistingOrAddNew(style);
        }

        void SetButtonStyle(IniFile.Section section, IFigure figure)
        {
            figure.Visible = !section.ReadBool("Hide", false) && section.ReadBool("Visible", true);

            var fontName = section["FontName"];
            double? fontSize = section.TryReadDouble("FontSize");
            bool fontUnderline = section.ReadBool("FontUnderline", false);
            var foreColor = section.ReadColor("ForeColor");
            var backColor = section.ReadColor("BackColor");

            TextStyle style = new TextStyle()
            {
                FontFamily = new FontFamily(fontName),
                Color = foreColor,
                Underline = fontUnderline
            };

            if (fontSize.HasValue)
            {
                style.FontSize = fontSize.Value * 1.7;
            }

            figure.Style = drawing.StyleManager.FindExistingOrAddNew(style);
        }

        IPoint GetPoint(IniFile.Section section, int index)
        {
            var pointIndex = GetPointIndex(section, index);
            EnsureSectionProcessed("Point" + pointIndex.ToString());
            var result = points[pointIndex];
            return result;
        }

        int GetPointIndex(IniFile.Section section, int index)
        {
            if (section["Points" + index.ToString()] != null)
            {
                return section.ReadInt("Points" + index.ToString());
            }
            return section.ReadInt("Point" + index.ToString());
        }

        IFigure GetParent(IniFile.Section section, int index)
        {
            var parentIndex = section.ReadInt("Parents" + index.ToString());
            var pointSectionTitle = "Figure" + parentIndex.ToString();
            EnsureSectionProcessed(pointSectionTitle);
            return figures[parentIndex];
        }

        double GetAuxInfo(IniFile.Section section, int index)
        {
            return section.ReadDouble("AuxInfo(" + index.ToString() + ")");
        }

        Point GetAuxPoint(IniFile.Section section, int index)
        {
            double x = section.ReadDouble("AuxPoints(" + index.ToString() + ").X");
            double y = section.ReadDouble("AuxPoints(" + index.ToString() + ").Y");
            return new Point(x, y);
        }

        void ProcessPoint(IniFile.Section section)
        {
            var type = section.ReadInt("Type");
            var parentFigure = section["ParentFigure"];
            if (!parentFigure.IsEmpty())
            {
                EnsureSectionProcessed("Figure" + parentFigure);
                return;
            }

            if (type == 8)
            {
                FindAndProcessArcWithPoint(section.GetTitleNumber("Point"));
                return;
            }

            if (type == 9)
            {
                FindAndProcessMidPoint(section.GetTitleNumber("Point"));
                return;
            }

            var x = section.ReadDouble("X");
            var y = section.ReadDouble("Y");
            FreePoint point = Factory.CreateFreePoint(drawing, new Point(x, y));

            point.Name = section["Name"];
            SetPointStyle(section, point);
            Actions.Add(drawing, point);
            int pointIndex = section.GetTitleNumber("Point");
            points[pointIndex] = point;
        }

        void FindAndProcessMidPoint(int pointIndex)
        {
            foreach (var section in sectionLookup.Values)
            {
                if (section["FigureType"] == "9")
                {
                    if (section.ReadInt("Points0") == pointIndex)
                    {
                        EnsureSectionProcessed(section.Title);
                        return;
                    }
                }
            }
        }

        void FindAndProcessArcWithPoint(int pointIndex)
        {
            foreach (var section in sectionLookup.Values)
            {
                if (section["FigureType"] == "8")
                {
                    if (section.ReadInt("Points5") == pointIndex
                        || section.ReadInt("Points6") == pointIndex)
                    {
                        EnsureSectionProcessed(section.Title);
                        return;
                    }
                }
            }
        }

        ShowHideControl[] buttons;
        PointBase[] points;
        IFigure[] figures;
        LabelBase[] labels;
        IFigure[] staticGraphics;

        void ProcessGeneral(IniFile.Section section)
        {
            var workArea = section["WorkArea"];
            if (workArea != null)
            {
                workArea = workArea.Trim('(', ')');
                var values = workArea.Split(',');
                var left = values[0];
                var bottom = values[1];
                var right = values[2];
                var top = values[3];
                var minX = double.Parse(left);
                var minY = double.Parse(bottom);
                var maxX = double.Parse(right);
                var maxY = double.Parse(top);
                drawing.CoordinateSystem.SetViewport(minX, maxX, minY, maxY);
            }

            var showAxes = section.ReadBool("ShowAxes", Settings.Instance.ShowGrid);
            var showGrid = section.ReadBool("ShowGrid", Settings.Instance.ShowGrid);

            drawing.CoordinateGrid.Visible = showAxes || showGrid;

            var pointCount = section.ReadInt("PointCount");
            points = new PointBase[pointCount + 1];

            var figureCount = section.ReadInt("FigureCount");
            figures = new IFigure[figureCount];

            var labelCount = section.ReadInt("LabelCount");
            labels = new LabelBase[labelCount + 1];

            var buttonCount = section.ReadInt("ButtonCount");
            buttons = new ShowHideControl[buttonCount + 1];

            var staticGraphicCount = section.ReadInt("StaticGraphicCount");
            staticGraphics = new IFigure[staticGraphicCount + 1];
        }
    }
}
