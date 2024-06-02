using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI.Selection;

namespace ConvertCAD
{
    [Transaction(TransactionMode.Manual)]
    public class ConvertCAD2Revit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document document = uidoc.Document; // 获取当前软件、文档

            Reference refer = uidoc.Selection.PickObject(
                ObjectType.Element,
                "选择CAD链接图层");   // 拾取CAD链接

            Element element = document.GetElement(refer);
            GeometryObject geoObj = element.GetGeometryObjectFromReference(refer);

            Category cadLinkCategory = null;    // 获取父Category

            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                if (document.GetElement(geoObj.GraphicsStyleId) is GraphicsStyle gs)
                {
                    cadLinkCategory = gs.GraphicsStyleCategory;
                }
            }

            if (cadLinkCategory == null)
            {
                return Result.Failed;   // 未匹配
            }

            Category wallsCategory = null;
            Category windowsCategory = null;
            Category doorsCategory = null;
            Category axesCategory = null;
            Category columnsCategory = null;

            List<GeometryObject> wallsObjs = new List<GeometryObject>();
            List<GeometryObject> windowsObjs = new List<GeometryObject>();
            List<GeometryObject> doorsObjs = new List<GeometryObject>();
            List<GeometryObject> axesObjs = new List<GeometryObject>();

            foreach (Category subCategory in cadLinkCategory.SubCategories)
            {
                // 匹配各图层名称
                if (("墙" == subCategory.Name) ||
                    ("WALL" == subCategory.Name) ||
                    ("Wall" == subCategory.Name) ||
                    ("wall" == subCategory.Name))
                {
                    wallsCategory = subCategory;
                }

                if (("门" == subCategory.Name) ||
                    ("DOOR" == subCategory.Name) ||
                    ("Door" == subCategory.Name) ||
                    ("door" == subCategory.Name))
                {
                    doorsCategory = subCategory;
                }

                if (("窗" == subCategory.Name) ||
                    ("WINDOW" == subCategory.Name) ||
                    ("Window" == subCategory.Name) ||
                    ("window" == subCategory.Name))
                {
                    windowsCategory = subCategory;
                }

                if (("轴线" == subCategory.Name) ||
                    ("AXIS" == subCategory.Name) ||
                    ("Axis" == subCategory.Name) ||
                    ("axis" == subCategory.Name))
                {
                    axesCategory = subCategory;
                }

                if (("柱" == subCategory.Name) ||
                    ("柱子" == subCategory.Name) ||
                    ("COLUMN" == subCategory.Name) ||
                    ("Column" == subCategory.Name) ||
                    ("column" == subCategory.Name))
                {
                    columnsCategory = subCategory;
                }
            }

            GeometryElement geometryElement = element.get_Geometry(new Options()); // 获取主图元

            Transaction trans = new Transaction(document, "生成柱");
            trans.Start();
            Level level = document.ActiveView.GenLevel;
            foreach (GeometryObject gObj in geometryElement)
            {
                //只取第一个元素
                GeometryInstance geomInstance = gObj as GeometryInstance;
                Transform totalTransform = geomInstance.Transform;

                foreach (var insObj in geomInstance.SymbolGeometry)
                {
                    Category insCategory = (document.GetElement(insObj.GraphicsStyleId) as GraphicsStyle).GraphicsStyleCategory;
                    if (wallsCategory != null && wallsCategory.Id == insCategory.Id)
                    {
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> endpoints = polyLine.GetCoordinates();
                            for (int i = 1; i < endpoints.Count; i++)
                            {
                                wallsObjs.Add(Line.CreateBound(
                                    totalTransform.OfPoint(endpoints[i - 1]),
                                    totalTransform.OfPoint(endpoints[i])
                                    ));
                            }
                        }
                        else
                        {
                            Line line = insObj as Line;
                            wallsObjs.Add(
                                Line.CreateBound(
                                    totalTransform.OfPoint(line.GetEndPoint(0)),
                                    totalTransform.OfPoint(line.GetEndPoint(1))
                                )
                            );
                        }
                    }
                    else if (windowsCategory != null && windowsCategory.Id == insCategory.Id)
                    {
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> endpoints = polyLine.GetCoordinates();
                            for (int i = 1; i < endpoints.Count; i++)
                            {
                                windowsObjs.Add(Line.CreateBound(
                                    totalTransform.OfPoint(endpoints[i - 1]),
                                    totalTransform.OfPoint(endpoints[i])
                                    ));
                            }
                        }
                        else
                        {
                            Line line = insObj as Line;
                            windowsObjs.Add(
                                Line.CreateBound(
                                    totalTransform.OfPoint(line.GetEndPoint(0)),
                                    totalTransform.OfPoint(line.GetEndPoint(1))
                                )
                            );
                        }
                    }
                    else if (doorsCategory != null && doorsCategory.Id == insCategory.Id)
                    {
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.GeometryInstance")
                        {
                            GeometryInstance geom = insObj as GeometryInstance;
                            Transform transform = geom.Transform;
                            List<XYZ> units = new List<XYZ>();
                            foreach (var unit in geom.SymbolGeometry)
                            {
                                if (unit.GetType().ToString() == "Autodesk.Revit.DB.Arc")
                                {
                                    Arc arc = unit as Arc;
                                    units.Add(totalTransform.OfPoint(transform.OfPoint(arc.GetEndPoint(0))));
                                    units.Add(totalTransform.OfPoint(transform.OfPoint(arc.GetEndPoint(1))));
                                }
                                else if (unit.GetType().ToString() == "Autodesk.Revit.DB.Line")
                                {
                                    Line line = unit as Line;
                                    units.Add(totalTransform.OfPoint(transform.OfPoint(line.GetEndPoint(0))));
                                    units.Add(totalTransform.OfPoint(transform.OfPoint(line.GetEndPoint(1))));
                                }
                            }

                            List<XYZ> isolatePoint = new List<XYZ>();
                            foreach (XYZ p1 in units)
                            {
                                int count = 0;
                                foreach (XYZ p2 in units)
                                {
                                    if (ElementPositionUtils.IsSameXYZ(p1, p2))
                                    {
                                        count++;
                                    }
                                }
                                if (1 == count)
                                {
                                    isolatePoint.Add(p1);
                                }
                            }
                            doorsObjs.Add(Line.CreateBound(isolatePoint[0], isolatePoint[1]));
                        }
                        else
                        {
                            doorsObjs.Add(insObj);
                        }
                    }
                    else if (axesCategory != null && axesCategory.Id == insCategory.Id)
                    {
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> endpoints = polyLine.GetCoordinates();
                            for (int i = 1; i < endpoints.Count; i++)
                            {
                                axesObjs.Add(Line.CreateBound(
                                    totalTransform.OfPoint(endpoints[i - 1]),
                                    totalTransform.OfPoint(endpoints[i])
                                    ));
                            }
                        }
                        else
                        {
                            Line line = insObj as Line;
                            axesObjs.Add(
                                Line.CreateBound(
                                    totalTransform.OfPoint(line.GetEndPoint(0)),
                                    totalTransform.OfPoint(line.GetEndPoint(1))
                                )
                            );
                        }
                    }
                    else if (columnsCategory != null && columnsCategory.Id == insCategory.Id)
                    {
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            XYZ pMax = polyLine.GetOutline().MaximumPoint;
                            XYZ pMin = polyLine.GetOutline().MinimumPoint;
                            double b = Math.Abs(pMin.X - pMax.X);
                            double h = Math.Abs(pMin.Y - pMax.Y);
                            XYZ pp = pMax.Add(pMin) / 2; // 柱中心点
                            pp = totalTransform.OfPoint(pp);

                            ElementPositionUtils.CreateColumn(document, pp, level, b, h);

                        }
                    }
                }
            }
            trans.Commit();

            List<WS> WSlist = new List<WS>();
            List<Line> Walloutlines = new List<Line>();

            foreach (GeometryObject gObj in wallsObjs)
            {
                Line wall = gObj as Line;
                Walloutlines.Add(wall);
            }


            for (int i = Walloutlines.Count - 1; i >= 0; i--)       //去除端线。
            {
                if ((Walloutlines[i].Length < 201 / 304.8) && !ElementPositionUtils.IsEqual(Walloutlines[i].Length, 100) || ElementPositionUtils.IsEqual(Walloutlines[i].Length, 250))
                {
                    Walloutlines.RemoveAt(i);
                }
            }


            ElementPositionUtils.Gen_WSlist(Walloutlines, ref WSlist);

            for (int i = WSlist.Count - 1; i >= 0; i--)
            {
                for (int j = i - 1; j >= 0; j--)
                {
                    if (ElementPositionUtils.IsSameWS(WSlist[i], WSlist[j]))
                    {
                        WSlist.RemoveAt(i);
                        break;
                    }
                }
            }


            for (int i = WSlist.Count - 1; i >= 0; i--)
            {
                if (!ElementPositionUtils.IsEqual(WSlist[i].BL1.Direction.X, 0) && !ElementPositionUtils.IsEqual(WSlist[i].BL1.Direction.Y, 0))
                {
                    WSlist.RemoveAt(i);
                }
            }

            /*
            string ans = "";
            foreach (WS ws in WSlist)
            {
                ans += ws.BL1.GetEndPoint(0).X.ToString() + ", " + ws.BL1.GetEndPoint(0).Y.ToString() + ", " + ws.BL1.GetEndPoint(0).Z.ToString() + ",,, " +
                    ws.BL1.GetEndPoint(1).X.ToString() + ", " + ws.BL1.GetEndPoint(1).Y.ToString() + ", " + ws.BL1.GetEndPoint(1).Z.ToString() + "\n" +
                    ws.BL2.GetEndPoint(0).X.ToString() + ", " + ws.BL2.GetEndPoint(0).Y.ToString() + ", " + ws.BL2.GetEndPoint(0).Z.ToString() + ",,, " +
                    ws.BL2.GetEndPoint(1).X.ToString() + ", " + ws.BL2.GetEndPoint(1).Y.ToString() + ", " + ws.BL2.GetEndPoint(1).Z.ToString() + "\n\n";
            }
            TaskDialog.Show("H", ans);
            */

            List<Point> connectpoints = new List<Point>();

            foreach (WS ws1 in WSlist)
            {
                ElementPositionUtils.Cal_Wallaxis(ws1, ref ws1.X, ref ws1.Y);
                foreach (WS ws2 in WSlist)
                {
                    if (ws1 != ws2 && ElementPositionUtils.Is_neighbor(ws1, ws2))
                    {
                        ElementPositionUtils.Cal_Wallaxis(ws2, ref ws2.X, ref ws2.Y);
                        Point p = Point.Create(new XYZ(ws1.X + ws2.X, ws1.Y + ws2.Y, 0));
                        int index = ElementPositionUtils.NotExistsIn(p, connectpoints);
                        if (index == connectpoints.Count)
                        {
                            connectpoints.Add(p);
                            ws1.KnotIndexs.Add(connectpoints.Count - 1);
                            ws2.KnotIndexs.Add(connectpoints.Count - 1);
                        }
                        else
                        {
                            bool isExist = false;
                            foreach (int i in ws1.KnotIndexs)
                            {
                                if (i == index)
                                {
                                    isExist = true;
                                }
                            }
                            if (!isExist) ws1.KnotIndexs.Add(index);

                            isExist = false;
                            foreach (int i in ws2.KnotIndexs)
                            {
                                if (i == index)
                                {
                                    isExist = true;
                                }
                            }
                            if (!isExist) ws2.KnotIndexs.Add(index);
                        }
                    }
                }
            }

            trans = new Transaction(document, "生成墙");
            trans.Start();
            foreach (WS ws in WSlist)
            {
                if (ws.KnotIndexs.Count == 0 || ws.KnotIndexs.Count == 1)
                {
                    var p1 = ws.BL1.GetEndPoint(0);
                    var p2 = ws.BL1.GetEndPoint(1);
                    var p3 = ws.BL2.GetEndPoint(0);
                    var p4 = ws.BL2.GetEndPoint(1);
                    Point point = null;
                    if (ElementPositionUtils.IsEqual(ws.BL1.Direction.X, 0))
                    {
                        if (ElementPositionUtils.IsEqual(p1.Y, p3.Y) || ElementPositionUtils.IsEqual(p1.Y, p4.Y))
                        {
                            point = Point.Create(new XYZ((p1.X + p3.X) / 2, p1.Y, 0));
                            connectpoints.Add(point);
                            ws.KnotIndexs.Add(connectpoints.Count - 1);
                        }
                        if (ElementPositionUtils.IsEqual(p2.Y, p3.Y) || ElementPositionUtils.IsEqual(p2.Y, p4.Y))
                        {
                            point = Point.Create(new XYZ((p1.X + p3.X) / 2, p2.Y, 0));
                            connectpoints.Add(point);
                            ws.KnotIndexs.Add(connectpoints.Count - 1);
                        }
                    }
                    else
                    {
                        if (ElementPositionUtils.IsEqual(p1.X, p3.X) || ElementPositionUtils.IsEqual(p1.X, p4.X))
                        {
                            point = Point.Create(new XYZ(p1.X, (p1.Y + p3.Y) / 2, 0));
                            connectpoints.Add(point);
                            ws.KnotIndexs.Add(connectpoints.Count - 1);
                        }
                        if (ElementPositionUtils.IsEqual(p2.X, p3.X) || ElementPositionUtils.IsEqual(p2.X, p4.X))
                        {
                            point = Point.Create(new XYZ(p2.X, (p1.Y + p3.Y) / 2, 0));
                            connectpoints.Add(point);
                            ws.KnotIndexs.Add(connectpoints.Count - 1);
                        }
                    }
                }

                List<Point> Knots = new List<Point>();
                double d = ElementPositionUtils.L2L_distance(ws.BL1, ws.BL2);
                foreach (int index in ws.KnotIndexs)
                {
                    Knots.Add(connectpoints[index]);
                    ElementPositionUtils.SortByCoord(ref Knots);
                    for (int j = 0; j < Knots.Count - 1; j++)
                    {
                        Line WallAxe = Line.CreateBound(Knots[j].Coord, Knots[j + 1].Coord);
                        if (WallAxe.Length > 126 / 304.8 && (ElementPositionUtils.IsEqual(WallAxe.Direction.X, 0) || ElementPositionUtils.IsEqual(WallAxe.Direction.Y, 0)))
                        {
                            ElementPositionUtils.CreateWall(document, WallAxe, level, d, 4000 / 304.8, 0);
                        }
                    }
                }
            }
            trans.Commit();

            List<Line> horizontalWindows = new List<Line>();
            List<Line> verticalWindows = new List<Line>();

            foreach (GeometryObject windowsObj in windowsObjs)
            {
                Line window = windowsObj as Line;
                if (ElementPositionUtils.IsEqual(window.GetEndPoint(0).X, window.GetEndPoint(1).X))
                {
                    verticalWindows.Add(window);
                }
                else if (ElementPositionUtils.IsEqual(window.GetEndPoint(0).Y, window.GetEndPoint(1).Y))
                {
                    horizontalWindows.Add(window);
                }
            }
            for (int i = 0; i < horizontalWindows.Count - 1; i++)
            {
                for (int j = 0; j < horizontalWindows.Count - 1 - i; j++)
                {
                    if (horizontalWindows[j].GetEndPoint(0).Y >
                        horizontalWindows[j + 1].GetEndPoint(1).Y)
                    {
                        var temp = horizontalWindows[j + 1];
                        horizontalWindows[j + 1] = horizontalWindows[j];
                        horizontalWindows[j] = temp;
                    }
                }
            }
            for (int i = 0; i < verticalWindows.Count - 1; i++)
            {
                for (int j = 0; j < verticalWindows.Count - 1 - i; j++)
                {
                    if (verticalWindows[j].GetEndPoint(0).X >
                        verticalWindows[j + 1].GetEndPoint(1).X)
                    {
                        var temp = verticalWindows[j + 1];
                        verticalWindows[j + 1] = verticalWindows[j];
                        verticalWindows[j] = temp;
                    }
                }
            }
            List<List<Line>> verticalSplitWindows = new List<List<Line>>();
            foreach (Line verticalWindow in verticalWindows)
            {
                bool is_exist = false;
                foreach (List<Line> verticalSplitWindow in verticalSplitWindows)
                {
                    if (verticalSplitWindow.Count > 0 &&
                       ElementPositionUtils.IsSameY(verticalWindow, verticalSplitWindow[0]) &&
                       ElementPositionUtils.Distance(verticalWindow, verticalSplitWindow[0]) < 500 / 304.8)
                    {
                        is_exist = true;
                        verticalSplitWindow.Add(verticalWindow);
                    }
                }
                if (!is_exist)
                {
                    List<Line> verticalSplitWindow = new List<Line>
                    {
                        verticalWindow
                    };
                    verticalSplitWindows.Add(verticalSplitWindow);
                }
            }
            List<List<Line>> horizontalSplitWindows = new List<List<Line>>();
            foreach (Line horizontalWindow in horizontalWindows)
            {
                bool is_exist = false;
                foreach (List<Line> horizontalSplitWindow in horizontalSplitWindows)
                {
                    if (horizontalSplitWindow.Count > 0 &&
                       ElementPositionUtils.IsSameX(horizontalWindow, horizontalSplitWindow[0]) &&
                       ElementPositionUtils.Distance(horizontalWindow, horizontalSplitWindow[0]) < 500 / 304.8)
                    {
                        is_exist = true;
                        horizontalSplitWindow.Add(horizontalWindow);
                    }
                }
                if (!is_exist)
                {
                    List<Line> horizontalSplitWindow = new List<Line>
                    {
                        horizontalWindow
                    };
                    horizontalSplitWindows.Add(horizontalSplitWindow);
                }
            }

            /*
            string ans = "";
            foreach(List<Line> horizontalSplitWindow in horizontalSplitWindows)
            {
                foreach(Line horizontalS in horizontalSplitWindow)
                {
                    ans += horizontalS.GetEndPoint(0).X.ToString() + "," +
                        horizontalS.GetEndPoint(0).Y.ToString() + "," +
                        horizontalS.GetEndPoint(0).Z.ToString() + ",,," +
                        horizontalS.GetEndPoint(1).X.ToString() + "," +
                        horizontalS.GetEndPoint(1).Y.ToString() + "," +
                        horizontalS.GetEndPoint(1).Z.ToString() + "\n";
                }
                ans += "\n";
            }
            TaskDialog.Show("H", ans);  // Debug
            */

            List<Point> endPoints = new List<Point>();
            trans = new Transaction(document, "生成窗");
            trans.Start();
            foreach (List<Line> horizontalSplitWindow in horizontalSplitWindows)
            {
                double width = 0;
                foreach (Line horizontalS in horizontalSplitWindow)
                {
                    endPoints.Add(Point.Create(horizontalS.GetEndPoint(0)));
                    endPoints.Add(Point.Create(horizontalS.GetEndPoint(1)));
                    width = horizontalS.Length;
                }
                double maxY = endPoints[0].Coord.Y;
                foreach (Point point in endPoints)
                {
                    if (point.Coord.Y > maxY)
                    {
                        maxY = point.Coord.Y;
                    }
                }
                double minY = endPoints[0].Coord.Y;
                foreach (Point point in endPoints)
                {
                    if (point.Coord.Y < minY)
                    {
                        minY = point.Coord.Y;
                    }
                }
                Point center_temp = ElementPositionUtils.GetCenter(endPoints);
                XYZ center = new XYZ(center_temp.Coord.X,
                                     center_temp.Coord.Y,
                                     level.ProjectElevation + 3);
                Line axesline = Line.CreateBound(
                    new XYZ(
                        horizontalSplitWindow[0].GetEndPoint(0).X,
                        center.Y,
                        horizontalSplitWindow[0].GetEndPoint(0).Z),
                    new XYZ(
                        horizontalSplitWindow[0].GetEndPoint(1).X,
                        center.Y,
                        horizontalSplitWindow[0].GetEndPoint(1).Z)
                    );
                Wall wall = ElementPositionUtils.CreateWall(document, axesline, level, maxY - minY, 4000 / 304.8, 0);
                ElementPositionUtils.CreateWindow(document, center, width, 3, wall, level);
                endPoints.Clear();
            }
            foreach (List<Line> verticalSplitWindow in verticalSplitWindows)
            {
                double width = 0;
                foreach (Line verticalS in verticalSplitWindow)
                {
                    endPoints.Add(Point.Create(verticalS.GetEndPoint(0)));
                    endPoints.Add(Point.Create(verticalS.GetEndPoint(1)));
                    width = verticalS.Length;
                }
                double maxX = endPoints[0].Coord.X;
                foreach (Point point in endPoints)
                {
                    if (point.Coord.X > maxX)
                    {
                        maxX = point.Coord.X;
                    }
                }
                double minX = endPoints[0].Coord.X;
                foreach (Point point in endPoints)
                {
                    if (point.Coord.X < minX)
                    {
                        minX = point.Coord.X;
                    }
                }
                Point center_temp = ElementPositionUtils.GetCenter(endPoints);
                XYZ center = new XYZ(center_temp.Coord.X,
                                     center_temp.Coord.Y,
                                     level.ProjectElevation + 3);
                Line axesline = Line.CreateBound(
                    new XYZ(
                        center.X,
                        verticalSplitWindow[0].GetEndPoint(0).Y,
                        verticalSplitWindow[0].GetEndPoint(0).Z),
                    new XYZ(
                        center.X,
                        verticalSplitWindow[0].GetEndPoint(1).Y,
                        verticalSplitWindow[0].GetEndPoint(1).Z)
                    );
                Wall wall = ElementPositionUtils.CreateWall(document, axesline, level, maxX - minX, 4000 / 304.8, 0);
                ElementPositionUtils.CreateWindow(document, center, width, 3, wall, level);
                endPoints.Clear();
            }
            trans.Commit();

            trans = new Transaction(document, "生成轴网");
            trans.Start();
            ElementPositionUtils.CreateAxes(document, axesObjs);
            trans.Commit();

            trans = new Transaction(document, "生成门");
            trans.Start();
            foreach (GeometryObject geo in doorsObjs)
            {
                double width = 0;
                Line door = geo as Line;
                XYZ doorStart = door.GetEndPoint(0);
                XYZ doorEnd = door.GetEndPoint(1);
                foreach (GeometryObject gObj in wallsObjs)
                {
                    Line wallLine = gObj as Line;
                    if (ElementPositionUtils.IsEqual(0, wallLine.Distance(doorStart)) ||
                        ElementPositionUtils.IsEqual(0, wallLine.Distance(doorEnd)))
                    {
                        width = (width == 0 ? wallLine.Length : Math.Min(width, wallLine.Length));
                    }
                }
                Wall wall = ElementPositionUtils.CreateWall(document, door, level, width, 4000 / 304.8, 0);
                ElementPositionUtils.CreateDoor(document, doorStart.Add(doorEnd) / 2, door.Length, 2200 / 304.8, wall, level);
            }
            trans.Commit();

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class ConvertWall2Revit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document document = uidoc.Document; // 获取当前软件、文档

            Reference refer = uidoc.Selection.PickObject(
                ObjectType.PointOnElement,
                "请拾取CAD链接中的墙");   // 拾取CAD链接

            Element element = document.GetElement(refer);
            GeometryElement geoElem = element.get_Geometry(new Options());
            GeometryObject geoObj = element.GetGeometryObjectFromReference(refer);

            Category targetCategory = null;    // 获取Category

            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                if (document.GetElement(geoObj.GraphicsStyleId) is GraphicsStyle gs)
                {
                    targetCategory = gs.GraphicsStyleCategory;
                }
            }

            if (targetCategory == null)
            {
                return Result.Failed;   // 未匹配
            }

            List<GeometryObject> wallsObjs = new List<GeometryObject>();

            foreach (GeometryObject gObj in geoElem)
            {
                //只取第一个元素
                GeometryInstance geomInstance = gObj as GeometryInstance;
                Transform transform = geomInstance.Transform;
                foreach (var insObj in geomInstance.SymbolGeometry)
                {
                    Category insCategory = (document.GetElement(insObj.GraphicsStyleId) as GraphicsStyle).GraphicsStyleCategory;
                    if (targetCategory.Id == insCategory.Id)
                    {
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> endpoints = polyLine.GetCoordinates();
                            for (int i = 1; i < endpoints.Count; i++)
                            {
                                wallsObjs.Add(Line.CreateBound(
                                    transform.OfPoint(endpoints[i - 1]),
                                    transform.OfPoint(endpoints[i])
                                    ));
                            }
                        }
                        else
                        {
                            Line line = insObj as Line;
                            wallsObjs.Add(
                                Line.CreateBound(
                                    transform.OfPoint(line.GetEndPoint(0)),
                                    transform.OfPoint(line.GetEndPoint(1))
                                )
                            );
                        }
                    }
                }
            }

            List<WS> WSlist = new List<WS>();
            List<Line> Walloutlines = new List<Line>();

            foreach (GeometryObject gObj in wallsObjs)
            {
                Line wall = gObj as Line;
                Walloutlines.Add(wall);
            }


            for (int i = Walloutlines.Count - 1; i >= 0; i--)       //去除端线。
            {
                if ((Walloutlines[i].Length < 201 / 304.8) && !ElementPositionUtils.IsEqual(Walloutlines[i].Length, 100) || ElementPositionUtils.IsEqual(Walloutlines[i].Length, 250))
                {
                    Walloutlines.RemoveAt(i);
                }
            }


            ElementPositionUtils.Gen_WSlist(Walloutlines, ref WSlist);

            for (int i = WSlist.Count - 1; i >= 0; i--)
            {
                for (int j = i - 1; j >= 0; j--)
                {
                    if (ElementPositionUtils.IsSameWS(WSlist[i], WSlist[j]))
                    {
                        WSlist.RemoveAt(i);
                        break;
                    }
                }
            }


            for (int i = WSlist.Count - 1; i >= 0; i--)
            {
                if (!ElementPositionUtils.IsEqual(WSlist[i].BL1.Direction.X, 0) && !ElementPositionUtils.IsEqual(WSlist[i].BL1.Direction.Y, 0))
                {
                    WSlist.RemoveAt(i);
                }
            }

            /*
            string ans = "";
            foreach (WS ws in WSlist)
            {
                ans += ws.BL1.GetEndPoint(0).X.ToString() + ", " + ws.BL1.GetEndPoint(0).Y.ToString() + ", " + ws.BL1.GetEndPoint(0).Z.ToString() + ",,, " +
                    ws.BL1.GetEndPoint(1).X.ToString() + ", " + ws.BL1.GetEndPoint(1).Y.ToString() + ", " + ws.BL1.GetEndPoint(1).Z.ToString() + "\n" +
                    ws.BL2.GetEndPoint(0).X.ToString() + ", " + ws.BL2.GetEndPoint(0).Y.ToString() + ", " + ws.BL2.GetEndPoint(0).Z.ToString() + ",,, " +
                    ws.BL2.GetEndPoint(1).X.ToString() + ", " + ws.BL2.GetEndPoint(1).Y.ToString() + ", " + ws.BL2.GetEndPoint(1).Z.ToString() + "\n\n";
            }
            TaskDialog.Show("H", ans);
            */

            List<Point> connectpoints = new List<Point>();

            foreach (WS ws1 in WSlist)
            {
                ElementPositionUtils.Cal_Wallaxis(ws1, ref ws1.X, ref ws1.Y);
                foreach (WS ws2 in WSlist)
                {
                    if (ws1 != ws2 && ElementPositionUtils.Is_neighbor(ws1, ws2))
                    {
                        ElementPositionUtils.Cal_Wallaxis(ws2, ref ws2.X, ref ws2.Y);
                        Point p = Point.Create(new XYZ(ws1.X + ws2.X, ws1.Y + ws2.Y, 0));
                        int index = ElementPositionUtils.NotExistsIn(p, connectpoints);
                        if (index == connectpoints.Count)
                        {
                            connectpoints.Add(p);
                            ws1.KnotIndexs.Add(connectpoints.Count - 1);
                            ws2.KnotIndexs.Add(connectpoints.Count - 1);
                        }
                        else
                        {
                            bool isExist = false;
                            foreach (int i in ws1.KnotIndexs)
                            {
                                if (i == index)
                                {
                                    isExist = true;
                                }
                            }
                            if (!isExist) ws1.KnotIndexs.Add(index);

                            isExist = false;
                            foreach (int i in ws2.KnotIndexs)
                            {
                                if (i == index)
                                {
                                    isExist = true;
                                }
                            }
                            if (!isExist) ws2.KnotIndexs.Add(index);
                        }
                    }
                }
            }




            //ElementPositionUtils.PrintLineList(WallAxis);
            Transaction trans = new Transaction(document, "生成墙");
            trans.Start();
            Level level = document.ActiveView.GenLevel;
            foreach (WS ws in WSlist)
            {
                if (ws.KnotIndexs.Count == 0 || ws.KnotIndexs.Count == 1)
                {
                    var p1 = ws.BL1.GetEndPoint(0);
                    var p2 = ws.BL1.GetEndPoint(1);
                    var p3 = ws.BL2.GetEndPoint(0);
                    var p4 = ws.BL2.GetEndPoint(1);
                    Point point = null;
                    if (ElementPositionUtils.IsEqual(ws.BL1.Direction.X, 0))
                    {
                        if (ElementPositionUtils.IsEqual(p1.Y, p3.Y) || ElementPositionUtils.IsEqual(p1.Y, p4.Y))
                        {
                            point = Point.Create(new XYZ((p1.X + p3.X) / 2, p1.Y, 0));
                            connectpoints.Add(point);
                            ws.KnotIndexs.Add(connectpoints.Count - 1);
                        }
                        if (ElementPositionUtils.IsEqual(p2.Y, p3.Y) || ElementPositionUtils.IsEqual(p2.Y, p4.Y))
                        {
                            point = Point.Create(new XYZ((p1.X + p3.X) / 2, p2.Y, 0));
                            connectpoints.Add(point);
                            ws.KnotIndexs.Add(connectpoints.Count - 1);
                        }
                    }
                    else
                    {
                        if (ElementPositionUtils.IsEqual(p1.X, p3.X) || ElementPositionUtils.IsEqual(p1.X, p4.X))
                        {
                            point = Point.Create(new XYZ(p1.X, (p1.Y + p3.Y) / 2, 0));
                            connectpoints.Add(point);
                            ws.KnotIndexs.Add(connectpoints.Count - 1);
                        }
                        if (ElementPositionUtils.IsEqual(p2.X, p3.X) || ElementPositionUtils.IsEqual(p2.X, p4.X))
                        {
                            point = Point.Create(new XYZ(p2.X, (p1.Y + p3.Y) / 2, 0));
                            connectpoints.Add(point);
                            ws.KnotIndexs.Add(connectpoints.Count - 1);
                        }
                    }
                }

                List<Point> Knots = new List<Point>();
                double d = ElementPositionUtils.L2L_distance(ws.BL1, ws.BL2);
                foreach (int index in ws.KnotIndexs)
                {
                    Knots.Add(connectpoints[index]);
                    ElementPositionUtils.SortByCoord(ref Knots);
                    for (int j = 0; j < Knots.Count - 1; j++)
                    {
                        Line WallAxe = Line.CreateBound(Knots[j].Coord, Knots[j + 1].Coord);
                        if (WallAxe.Length > 126 / 304.8 && (ElementPositionUtils.IsEqual(WallAxe.Direction.X, 0) || ElementPositionUtils.IsEqual(WallAxe.Direction.Y, 0)))
                        {
                            ElementPositionUtils.CreateWall(document, WallAxe, level, d, 4000 / 304.8, 0);
                        }
                    }
                }
            }
            trans.Commit();
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class ConvertWindow2Revit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document document = uidoc.Document; // 获取当前软件、文档

            Reference refer = uidoc.Selection.PickObject(
                ObjectType.PointOnElement,
                "请拾取CAD链接中的窗户");   // 拾取CAD链接

            Element element = document.GetElement(refer);
            GeometryElement geoElem = element.get_Geometry(new Options());
            GeometryObject geoObj = element.GetGeometryObjectFromReference(refer);

            Category targetCategory = null;    // 获取Category

            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                if (document.GetElement(geoObj.GraphicsStyleId) is GraphicsStyle gs)
                {
                    targetCategory = gs.GraphicsStyleCategory;
                }
            }

            if (targetCategory == null)
            {
                return Result.Failed;   // 未匹配
            }

            List<GeometryObject> windowsObjs = new List<GeometryObject>();

            foreach (GeometryObject gObj in geoElem)
            {
                //只取第一个元素
                GeometryInstance geomInstance = gObj as GeometryInstance;
                Transform transform = geomInstance.Transform;
                foreach (var insObj in geomInstance.SymbolGeometry)
                {
                    Category insCategory = (document.GetElement(insObj.GraphicsStyleId) as GraphicsStyle).GraphicsStyleCategory;
                    if (targetCategory.Id == insCategory.Id)
                    {
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> endpoints = polyLine.GetCoordinates();
                            for (int i = 1; i < endpoints.Count; i++)
                            {
                                windowsObjs.Add(Line.CreateBound(
                                    transform.OfPoint(endpoints[i - 1]),
                                    transform.OfPoint(endpoints[i])
                                    ));
                            }
                        }
                        else
                        {
                            Line line = insObj as Line;
                            windowsObjs.Add(
                                Line.CreateBound(
                                    transform.OfPoint(line.GetEndPoint(0)),
                                    transform.OfPoint(line.GetEndPoint(1))
                                )
                            );
                        }
                    }
                }
            }
            List<Line> horizontalWindows = new List<Line>();
            List<Line> verticalWindows = new List<Line>();

            foreach (GeometryObject windowsObj in windowsObjs)
            {
                Line window = windowsObj as Line;
                if (ElementPositionUtils.IsEqual(window.GetEndPoint(0).X, window.GetEndPoint(1).X))
                {
                    verticalWindows.Add(window);
                }
                else if (ElementPositionUtils.IsEqual(window.GetEndPoint(0).Y, window.GetEndPoint(1).Y))
                {
                    horizontalWindows.Add(window);
                }
            }
            for (int i = 0; i < horizontalWindows.Count - 1; i++)
            {
                for (int j = 0; j < horizontalWindows.Count - 1 - i; j++)
                {
                    if (horizontalWindows[j].GetEndPoint(0).Y >
                        horizontalWindows[j + 1].GetEndPoint(1).Y)
                    {
                        var temp = horizontalWindows[j + 1];
                        horizontalWindows[j + 1] = horizontalWindows[j];
                        horizontalWindows[j] = temp;
                    }
                }
            }
            for (int i = 0; i < verticalWindows.Count - 1; i++)
            {
                for (int j = 0; j < verticalWindows.Count - 1 - i; j++)
                {
                    if (verticalWindows[j].GetEndPoint(0).X >
                        verticalWindows[j + 1].GetEndPoint(1).X)
                    {
                        var temp = verticalWindows[j + 1];
                        verticalWindows[j + 1] = verticalWindows[j];
                        verticalWindows[j] = temp;
                    }
                }
            }
            List<List<Line>> verticalSplitWindows = new List<List<Line>>();
            foreach (Line verticalWindow in verticalWindows)
            {
                bool is_exist = false;
                foreach (List<Line> verticalSplitWindow in verticalSplitWindows)
                {
                    if (verticalSplitWindow.Count > 0 &&
                       ElementPositionUtils.IsSameY(verticalWindow, verticalSplitWindow[0]) &&
                       ElementPositionUtils.Distance(verticalWindow, verticalSplitWindow[0]) < 500 / 304.8)
                    {
                        is_exist = true;
                        verticalSplitWindow.Add(verticalWindow);
                    }
                }
                if (!is_exist)
                {
                    List<Line> verticalSplitWindow = new List<Line>
                    {
                        verticalWindow
                    };
                    verticalSplitWindows.Add(verticalSplitWindow);
                }
            }
            List<List<Line>> horizontalSplitWindows = new List<List<Line>>();
            foreach (Line horizontalWindow in horizontalWindows)
            {
                bool is_exist = false;
                foreach (List<Line> horizontalSplitWindow in horizontalSplitWindows)
                {
                    if (horizontalSplitWindow.Count > 0 &&
                       ElementPositionUtils.IsSameX(horizontalWindow, horizontalSplitWindow[0]) &&
                       ElementPositionUtils.Distance(horizontalWindow, horizontalSplitWindow[0]) < 500 / 304.8)
                    {
                        is_exist = true;
                        horizontalSplitWindow.Add(horizontalWindow);
                    }
                }
                if (!is_exist)
                {
                    List<Line> horizontalSplitWindow = new List<Line>
                    {
                        horizontalWindow
                    };
                    horizontalSplitWindows.Add(horizontalSplitWindow);
                }
            }

            /*
            string ans = "";
            foreach(List<Line> horizontalSplitWindow in verticalSplitWindows)
            {
                ans += "n\n";
                foreach(Line horizontalS in horizontalSplitWindow)
                {
                    ans += horizontalS.GetEndPoint(0).X.ToString() + "," +
                        horizontalS.GetEndPoint(0).Y.ToString() + "," +
                        horizontalS.GetEndPoint(0).Z.ToString() + ",,," +
                        horizontalS.GetEndPoint(1).X.ToString() + "," +
                        horizontalS.GetEndPoint(1).Y.ToString() + "," +
                        horizontalS.GetEndPoint(1).Z.ToString() + "\n";
                }
                ans += "\n";
            }
            TaskDialog.Show("H", ans);  // Debug
            */
            

            List<Point> endPoints = new List<Point>();
            Transaction trans = new Transaction(document, "生成窗");
            trans.Start();
            Level level = document.ActiveView.GenLevel;
            foreach (List<Line> horizontalSplitWindow in horizontalSplitWindows)
            {
                double width = 0;
                foreach (Line horizontalS in horizontalSplitWindow)
                {
                    endPoints.Add(Point.Create(horizontalS.GetEndPoint(0)));
                    endPoints.Add(Point.Create(horizontalS.GetEndPoint(1)));
                    width = horizontalS.Length;
                }
                double maxY = endPoints[0].Coord.Y;
                foreach (Point point in endPoints)
                {
                    if (point.Coord.Y > maxY)
                    {
                        maxY = point.Coord.Y;
                    }
                }
                double minY = endPoints[0].Coord.Y;
                foreach (Point point in endPoints)
                {
                    if (point.Coord.Y < minY)
                    {
                        minY = point.Coord.Y;
                    }
                }
                Point center_temp = ElementPositionUtils.GetCenter(endPoints);
                XYZ center = new XYZ(center_temp.Coord.X,
                                     center_temp.Coord.Y,
                                     level.ProjectElevation + 3);
                Line axesline = Line.CreateBound(
                    new XYZ(
                        horizontalSplitWindow[0].GetEndPoint(0).X,
                        center.Y,
                        horizontalSplitWindow[0].GetEndPoint(0).Z),
                    new XYZ(
                        horizontalSplitWindow[0].GetEndPoint(1).X,
                        center.Y,
                        horizontalSplitWindow[0].GetEndPoint(1).Z)
                    );
                Wall wall = ElementPositionUtils.CreateWall(document, axesline, level, maxY - minY, 4000 / 304.8, 0);
                ElementPositionUtils.CreateWindow(document, center, width, 3600 / 304.8, wall, level);
                endPoints.Clear();
            }
            foreach (List<Line> verticalSplitWindow in verticalSplitWindows)
            {
                double width = 0;
                foreach (Line verticalS in verticalSplitWindow)
                {
                    endPoints.Add(Point.Create(verticalS.GetEndPoint(0)));
                    endPoints.Add(Point.Create(verticalS.GetEndPoint(1)));
                    width = verticalS.Length;
                }
                double maxX = endPoints[0].Coord.X;
                foreach (Point point in endPoints)
                {
                    if (point.Coord.X > maxX)
                    {
                        maxX = point.Coord.X;
                    }
                }
                double minX = endPoints[0].Coord.X;
                foreach (Point point in endPoints)
                {
                    if (point.Coord.X < minX)
                    {
                        minX = point.Coord.X;
                    }
                }
                Point center_temp = ElementPositionUtils.GetCenter(endPoints);
                XYZ center = new XYZ(center_temp.Coord.X,
                                     center_temp.Coord.Y,
                                     level.ProjectElevation + 3);
                Line axesline = Line.CreateBound(
                    new XYZ(
                        center.X,
                        verticalSplitWindow[0].GetEndPoint(0).Y,
                        verticalSplitWindow[0].GetEndPoint(0).Z),
                    new XYZ(
                        center.X,
                        verticalSplitWindow[0].GetEndPoint(1).Y,
                        verticalSplitWindow[0].GetEndPoint(1).Z)
                    );
                Wall wall = ElementPositionUtils.CreateWall(document, axesline, level, maxX - minX, 4000 / 304.8, 0);
                ElementPositionUtils.CreateWindow(document, center, width, 3600 / 304.8, wall, level);
                endPoints.Clear();
            }
            trans.Commit();
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class ConvertAxis2Revit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document document = uidoc.Document; // 获取当前软件、文档

            Reference refer = uidoc.Selection.PickObject(
                ObjectType.PointOnElement,
                "请拾取CAD链接中的轴线");   // 拾取CAD链接

            Element element = document.GetElement(refer);
            GeometryElement geoElem = element.get_Geometry(new Options());
            GeometryObject geoObj = element.GetGeometryObjectFromReference(refer);

            Category targetCategory = null;    // 获取Category

            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                if (document.GetElement(geoObj.GraphicsStyleId) is GraphicsStyle gs)
                {
                    targetCategory = gs.GraphicsStyleCategory;
                }
            }

            if (targetCategory == null)
            {
                return Result.Failed;   // 未匹配
            }

            List<GeometryObject> axesObjs = new List<GeometryObject>();

            foreach (GeometryObject gObj in geoElem)
            {
                //只取第一个元素
                GeometryInstance geomInstance = gObj as GeometryInstance;
                Transform transform = geomInstance.Transform;
                foreach (var insObj in geomInstance.SymbolGeometry)
                {
                    Category insCategory = (document.GetElement(insObj.GraphicsStyleId) as GraphicsStyle).GraphicsStyleCategory;
                    if (targetCategory.Id == insCategory.Id)
                    {
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> endpoints = polyLine.GetCoordinates();
                            for (int i = 1; i < endpoints.Count; i++)
                            {
                                axesObjs.Add(Line.CreateBound(
                                    transform.OfPoint(endpoints[i - 1]),
                                    transform.OfPoint(endpoints[i])
                                    ));
                            }
                        }
                        else
                        {
                            Line line = insObj as Line;
                            axesObjs.Add(
                                Line.CreateBound(
                                    transform.OfPoint(line.GetEndPoint(0)),
                                    transform.OfPoint(line.GetEndPoint(1))
                                )
                            );
                        }
                    }
                }
            }

            Transaction trans = new Transaction(document, "生成轴网");
            trans.Start();
            ElementPositionUtils.CreateAxes(document, axesObjs);
            trans.Commit();
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class ConvertColumn2Revit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document document = uidoc.Document; // 获取当前软件、文档

            Reference refer = uidoc.Selection.PickObject(
                ObjectType.PointOnElement,
                "请拾取CAD链接中的柱");   // 拾取CAD链接

            Element element = document.GetElement(refer);
            GeometryElement geoElem = element.get_Geometry(new Options());
            GeometryObject geoObj = element.GetGeometryObjectFromReference(refer);

            Category targetCategory = null;    // 获取Category

            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                if (document.GetElement(geoObj.GraphicsStyleId) is GraphicsStyle gs)
                {
                    targetCategory = gs.GraphicsStyleCategory;
                }
            }

            if (targetCategory == null)
            {
                return Result.Failed;   // 未匹配
            }

            Transaction trans = new Transaction(document, "创建柱");
            trans.Start();
            Level level = document.ActiveView.GenLevel;
            foreach (GeometryObject gObj in geoElem)
            {
                //只取第一个元素
                GeometryInstance geomInstance = gObj as GeometryInstance;
                Transform transform = geomInstance.Transform;
                foreach (var insObj in geomInstance.SymbolGeometry)
                {
                    Category insCategory = (document.GetElement(insObj.GraphicsStyleId) as GraphicsStyle).GraphicsStyleCategory;
                    if (targetCategory.Id == insCategory.Id)
                    {
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            XYZ pMax = polyLine.GetOutline().MaximumPoint;
                            XYZ pMin = polyLine.GetOutline().MinimumPoint;
                            double b = Math.Abs(pMin.X - pMax.X);
                            double h = Math.Abs(pMin.Y - pMax.Y);
                            XYZ pp = pMax.Add(pMin) / 2; // 柱中心点
                            pp = transform.OfPoint(pp);
                            ElementPositionUtils.CreateColumn(document, pp, level, b, h);
                        }
                    }
                }
            }
            trans.Commit();
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class ConvertDoor2Revit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document document = uidoc.Document; // 获取当前软件、文档

            Reference refer = uidoc.Selection.PickObject(
                ObjectType.PointOnElement,
                "请拾取CAD链接中的门");   // 拾取CAD链接

            Element element = document.GetElement(refer);
            GeometryElement geoElem = element.get_Geometry(new Options());
            GeometryObject geoObj = element.GetGeometryObjectFromReference(refer);

            Category targetCategory = null;    // 获取Category

            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                if (document.GetElement(geoObj.GraphicsStyleId) is GraphicsStyle gs)
                {
                    targetCategory = gs.GraphicsStyleCategory;
                }
            }

            if (targetCategory == null)
            {
                return Result.Failed;   // 未匹配
            }

            List<GeometryObject> doorsObjs = new List<GeometryObject>();

            foreach (GeometryObject gObj in geoElem)
            {
                //只取第一个元素
                GeometryInstance geomInstance = gObj as GeometryInstance;
                Transform totalTransform = geomInstance.Transform;
                foreach (var insObj in geomInstance.SymbolGeometry)
                {
                    Category insCategory = (document.GetElement(insObj.GraphicsStyleId) as GraphicsStyle).GraphicsStyleCategory;
                    if (targetCategory.Id == insCategory.Id)
                    {
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.GeometryInstance")
                        {
                            GeometryInstance geom = insObj as GeometryInstance;
                            Transform transform = geom.Transform;
                            List<XYZ> units = new List<XYZ>();
                            foreach (var unit in geom.SymbolGeometry)
                            {
                                if (unit.GetType().ToString() == "Autodesk.Revit.DB.Arc")
                                {
                                    Arc arc = unit as Arc;
                                    units.Add(totalTransform.OfPoint(transform.OfPoint(arc.GetEndPoint(0))));
                                    units.Add(totalTransform.OfPoint(transform.OfPoint(arc.GetEndPoint(1))));
                                }
                                else if (unit.GetType().ToString() == "Autodesk.Revit.DB.Line")
                                {
                                    Line line = unit as Line;
                                    units.Add(totalTransform.OfPoint(transform.OfPoint(line.GetEndPoint(0))));
                                    units.Add(totalTransform.OfPoint(transform.OfPoint(line.GetEndPoint(1))));
                                }
                            }

                            List<XYZ> isolatePoint = new List<XYZ>();
                            foreach (XYZ p1 in units)
                            {
                                int count = 0;
                                foreach (XYZ p2 in units)
                                {
                                    if (ElementPositionUtils.IsSameXYZ(p1, p2))
                                    {
                                        count++;
                                    }
                                }
                                if (1 == count)
                                {
                                    isolatePoint.Add(p1);
                                }
                            }
                            doorsObjs.Add(Line.CreateBound(isolatePoint[0], isolatePoint[1]));
                        }
                        else
                        {
                            doorsObjs.Add(insObj);
                        }
                    }
                }
            }

            refer = uidoc.Selection.PickObject(
                ObjectType.PointOnElement,
                "请拾取CAD链接中的墙");   // 拾取CAD链接

            element = document.GetElement(refer);
            geoElem = element.get_Geometry(new Options());
            geoObj = element.GetGeometryObjectFromReference(refer);

            targetCategory = null;    // 获取Category

            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                if (document.GetElement(geoObj.GraphicsStyleId) is GraphicsStyle gs)
                {
                    targetCategory = gs.GraphicsStyleCategory;
                }
            }

            if (targetCategory == null)
            {
                return Result.Failed;   // 未匹配
            }

            List<GeometryObject> wallsObjs = new List<GeometryObject>();

            foreach (GeometryObject gObj in geoElem)
            {
                //只取第一个元素
                GeometryInstance geomInstance = gObj as GeometryInstance;
                Transform transform = geomInstance.Transform;
                foreach (var insObj in geomInstance.SymbolGeometry)
                {
                    Category insCategory = (document.GetElement(insObj.GraphicsStyleId) as GraphicsStyle).GraphicsStyleCategory;
                    if (targetCategory.Id == insCategory.Id)
                    {
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> endpoints = polyLine.GetCoordinates();
                            for (int i = 1; i < endpoints.Count; i++)
                            {
                                wallsObjs.Add(Line.CreateBound(
                                    transform.OfPoint(endpoints[i - 1]),
                                    transform.OfPoint(endpoints[i])
                                    ));
                            }
                        }
                        else
                        {
                            Line line = insObj as Line;
                            wallsObjs.Add(
                                Line.CreateBound(
                                    transform.OfPoint(line.GetEndPoint(0)),
                                    transform.OfPoint(line.GetEndPoint(1))
                                )
                            );
                        }
                    }
                }
            }

            Transaction trans = new Transaction(document, "生成门");
            trans.Start();
            Level level = document.ActiveView.GenLevel;
            foreach (GeometryObject geo in doorsObjs)
            {
                double width = 0;
                Line door = geo as Line;
                XYZ doorStart = door.GetEndPoint(0);
                XYZ doorEnd = door.GetEndPoint(1);
                foreach (GeometryObject gObj in wallsObjs)
                {
                    Line wallLine = gObj as Line;
                    if (ElementPositionUtils.IsEqual(0, wallLine.Distance(doorStart)) ||
                        ElementPositionUtils.IsEqual(0, wallLine.Distance(doorEnd)))
                    {
                        width = (width == 0 ? wallLine.Length : Math.Min(width, wallLine.Length));
                    }
                }
                Wall wall = ElementPositionUtils.CreateWall(document, door, level, width, 4000 / 304.8, 0);
                ElementPositionUtils.CreateDoor(document, doorStart.Add(doorEnd) / 2, door.Length, 2200 / 304.8, wall, level);
            }
            trans.Commit();
            return Result.Succeeded;
        }
    }

    public class ElementPositionUtils
    {
        public static bool IsEqual(double x, double y)
        {
            return Math.Abs(x - y) < 1e-3;
        }

        public static bool IsSameX(Line l1, Line l2)
        {
            return ((Math.Abs(l1.GetEndPoint(0).X - l2.GetEndPoint(0).X) < 1e-3 &&
                     Math.Abs(l1.GetEndPoint(1).X - l2.GetEndPoint(1).X) < 1e-3) ||
                    (Math.Abs(l1.GetEndPoint(0).X - l2.GetEndPoint(1).X) < 1e-3 &&
                     Math.Abs(l1.GetEndPoint(1).X - l2.GetEndPoint(0).X) < 1e-3));
        }

        public static bool IsSameY(Line l1, Line l2)
        {
            return ((Math.Abs(l1.GetEndPoint(0).Y - l2.GetEndPoint(0).Y) < 1e-3 &&
                     Math.Abs(l1.GetEndPoint(1).Y - l2.GetEndPoint(1).Y) < 1e-3) ||
                    (Math.Abs(l1.GetEndPoint(0).Y - l2.GetEndPoint(1).Y) < 1e-3 &&
                     Math.Abs(l1.GetEndPoint(1).Y - l2.GetEndPoint(0).Y) < 1e-3));
        }

        public static bool IsSameXYZ(XYZ p1, XYZ p2)
        {
            return (IsEqual(p1.X, p2.X) && IsEqual(p1.Y, p2.Y) && IsEqual(p1.Z, p2.Z));
        }

        public static void PrintPointList(List<Point> points)
        {
            string ans = "";
            foreach (Point point in points)
            {
                ans += point.Coord.X.ToString() + " " +
                                     point.Coord.Y.ToString() + " " +
                                     point.Coord.Z.ToString() + "\n";
            }
            TaskDialog.Show("Point", ans);
        }

        public static void PrintLineList(List<Line> lines)
        {
            string ans = "";
            foreach (Line line in lines)
            {
                ans += line.GetEndPoint(0).X.ToString() + " " +
                       line.GetEndPoint(0).Y.ToString() + " " +
                       line.GetEndPoint(0).Z.ToString() + ", " +
                       line.GetEndPoint(1).X.ToString() + " " +
                       line.GetEndPoint(1).Y.ToString() + " " +
                       line.GetEndPoint(1).Z.ToString() + "\n";
            }
            TaskDialog.Show("Line", ans);
        }

        public static Point GetCenter(List<Point> pointlist)
        {
            double X_sum = 0, Y_sum = 0, Z_sum = 0;
            for (int i = 0; i < pointlist.Count; i++)
            {
                X_sum += pointlist[i].Coord.X;
                Y_sum += pointlist[i].Coord.Y;
                Z_sum += pointlist[i].Coord.Z;
            }

            double center_X = X_sum / pointlist.Count;
            double center_Y = Y_sum / pointlist.Count;
            double center_Z = Z_sum / pointlist.Count;
            Point center = Point.Create(new XYZ(center_X, center_Y, center_Z));
            return center;
        }

        public static double Distance(Line l1, Line l2)
        {
            return l1.Distance(l2.GetEndPoint(0));
        }

        public static Wall CreateWall(Document document, Line axisLine, Level level, double width, double height, double offset)
        {
            FilteredElementCollector filter = new FilteredElementCollector(document);
            filter.OfClass(typeof(WallType));
            string attr = "常规 - " + Math.Round(width * 304.8, 3).ToString() + "mm";
            List<WallType> familySymbols = new List<WallType>();
            foreach (WallType familySymbol in filter)
            {
                if (familySymbol.GetParameters("族名称")[0].AsString() == "基本墙")
                {
                    familySymbols.Add(familySymbol);
                }
            }

            int i;
            bool bo = false;
            int j = 0;
            for (i = 0; i < familySymbols.Count; i++)
            {
                if (attr == familySymbols[i].Name)
                {
                    bo = true;
                    j = i;
                }
            }
            if (bo == true)
            {
                return Wall.Create(document, axisLine, familySymbols[j].Id, level.Id, height, offset, false, false);
            }
            else
            {
                for (i = 0; i < familySymbols.Count; i++)
                {
                    if ("常规 - 200mm" == familySymbols[i].Name) { break; }
                }
                WallType fam = familySymbols[i];
                WallType coluType = fam.Duplicate(attr) as WallType;
                CompoundStructure cs = coluType.GetCompoundStructure();
                int layerIndex = cs.GetFirstCoreLayerIndex();
                cs.SetLayerWidth(layerIndex, width);
                coluType.SetCompoundStructure(cs);
                return Wall.Create(document, axisLine, coluType.Id, level.Id, height, offset, false, false);
            }
        }

        public static void CreateColumn(Document document, XYZ center, Level level, double b, double h)
        {
            FilteredElementCollector filter = new FilteredElementCollector(document);
            filter.OfClass(typeof(FamilySymbol));
            string bh = Math.Round(b * 304.8, 3).ToString() + " " + "x" + " " + Math.Round(h * 304.8, 3).ToString() + "mm";
            List<FamilySymbol> familySymbols = new List<FamilySymbol>();
            foreach (FamilySymbol familySymbol in filter)
            {
                if (familySymbol.GetParameters("族名称")[0].AsString() == "混凝土 - 矩形 - 柱")
                {
                    familySymbols.Add(familySymbol);
                }
            }

            int i;
            bool bo = false;
            int j = 0;
            for (i = 0; i < familySymbols.Count; i++)
            {
                if (bh == familySymbols[i].Name)
                {
                    bo = true;
                    j = i;
                }
            }
            if (bo == true)
            {
                document.Create.NewFamilyInstance(center, familySymbols[j], level, StructuralType.Column);
            }
            else
            {
                FamilySymbol fam = familySymbols[0];
                ElementType coluType = fam.Duplicate(bh);
                coluType.GetParameters("b")[0].Set(b);
                coluType.GetParameters("h")[0].Set(h);
                FamilySymbol fs = coluType as FamilySymbol;
                document.Create.NewFamilyInstance(center, fs, level, StructuralType.Column);
            }
        }

        public static void CreateAxes(Document document, List<GeometryObject> axesObjs)
        {
            List<Line> horizontalAxis = new List<Line>();
            List<Line> verticalAxis = new List<Line>();

            foreach (GeometryObject axesObj in axesObjs)
            {
                Line axis = axesObj as Line;
                if (IsEqual(axis.GetEndPoint(0).X, axis.GetEndPoint(1).X))
                {
                    verticalAxis.Add(axis);
                }
                else if (IsEqual(axis.GetEndPoint(0).Y, axis.GetEndPoint(1).Y))
                {
                    horizontalAxis.Add(axis);
                }
            }
            for (int i = 0; i < horizontalAxis.Count - 1; i++)
            {
                for (int j = 0; j < horizontalAxis.Count - 1 - i; j++)
                {
                    if (horizontalAxis[j].GetEndPoint(0).Y >
                        horizontalAxis[j + 1].GetEndPoint(1).Y)
                    {
                        var temp = horizontalAxis[j + 1];
                        horizontalAxis[j + 1] = horizontalAxis[j];
                        horizontalAxis[j] = temp;
                    }
                }
            }
            for (int i = 0; i < verticalAxis.Count - 1; i++)
            {
                for (int j = 0; j < verticalAxis.Count - 1 - i; j++)
                {
                    if (verticalAxis[j].GetEndPoint(0).X >
                        verticalAxis[j + 1].GetEndPoint(1).X)
                    {
                        var temp = verticalAxis[j + 1];
                        verticalAxis[j + 1] = verticalAxis[j];
                        verticalAxis[j] = temp;
                    }
                }
            }
            if (verticalAxis.Count > 0)
            {
                Grid vertical = Grid.Create(document, verticalAxis[0]);
                vertical.Name = "1";
                for (int i = 1; i < verticalAxis.Count; i++)
                {
                    Grid.Create(document, verticalAxis[i]);
                }
            }
            if (horizontalAxis.Count > 0)
            {
                Grid horizontal = Grid.Create(document, horizontalAxis[0]);
                horizontal.Name = "A";
                for (int i = 1; i < horizontalAxis.Count; i++)
                {
                    Grid.Create(document, horizontalAxis[i]);
                }
            }
        }

        public static void CreateWindow(Document document, XYZ center, double width, double height, Wall desWall, Level level)
        {
            FilteredElementCollector filter = new FilteredElementCollector(document);
            filter.OfClass(typeof(FamilySymbol));
            string bh = Math.Round(width * 304.8, 3).ToString() + " " + "x" + " " +
                        Math.Round(height * 304.8, 3).ToString() + "mm";
            List<FamilySymbol> familySymbols = new List<FamilySymbol>();
            foreach (FamilySymbol familySymbol in filter)
            {
                if (familySymbol.GetParameters("族名称")[0].AsString() == "固定")
                {
                    familySymbols.Add(familySymbol);
                }
            }

            int i;
            bool bo = false;
            int j = 0;
            for (i = 0; i < familySymbols.Count; i++)
            {
                if (bh == familySymbols[i].Name)
                {
                    bo = true;
                    j = i;
                }
            }
            if (bo == true)
            {
                document.Create.NewFamilyInstance(center, familySymbols[j], desWall, level, StructuralType.NonStructural);
            }
            else
            {
                FamilySymbol fam = familySymbols[0];
                ElementType coluType = fam.Duplicate(bh);
                coluType.GetParameters("宽度")[0].Set(width);
                coluType.GetParameters("高度")[0].Set(height);
                FamilySymbol fs = coluType as FamilySymbol;
                document.Create.NewFamilyInstance(center, fs, desWall, level, StructuralType.NonStructural);
            }
        }

        public static void CreateDoor(Document document, XYZ center, double width, double height, Wall desWall, Level level)
        {
            FilteredElementCollector filter = new FilteredElementCollector(document);
            filter.OfClass(typeof(FamilySymbol));
            string bh = Math.Round(width * 304.8, 3).ToString() + " " + "x" + " " +
                        Math.Round(height * 304.8, 3).ToString() + "mm";
            List<FamilySymbol> familySymbols = new List<FamilySymbol>();
            foreach (FamilySymbol familySymbol in filter)
            {
                if (familySymbol.GetParameters("族名称")[0].AsString() == "单扇 - 与墙齐")
                {
                    familySymbols.Add(familySymbol);
                }
            }

            int i;
            bool bo = false;
            int j = 0;
            for (i = 0; i < familySymbols.Count; i++)
            {
                if (bh == familySymbols[i].Name)
                {
                    bo = true;
                    j = i;
                }
            }
            if (bo == true)
            {
                document.Create.NewFamilyInstance(center, familySymbols[j], desWall, level, StructuralType.NonStructural);
            }
            else
            {
                FamilySymbol fam = familySymbols[0];
                ElementType coluType = fam.Duplicate(bh);
                coluType.GetParameters("宽度")[0].Set(width);
                coluType.GetParameters("高度")[0].Set(height);
                FamilySymbol fs = coluType as FamilySymbol;
                document.Create.NewFamilyInstance(center, fs, desWall, level, StructuralType.NonStructural);
            }
        }

        public static double P2P_distance(Point p1, Point p2)
        {
            return Math.Sqrt(Math.Pow(p1.Coord.X - p2.Coord.X, 2) + Math.Pow(p1.Coord.Y - p2.Coord.Y, 2));
        }

        public static double L2L_distance(Line l1, Line l2)
        {
            if (IsEqual(l1.Direction.X, 0))
            {
                return Math.Abs(l1.GetEndPoint(0).X - l2.GetEndPoint(0).X);
            }
            else
            {
                return Math.Abs(l1.GetEndPoint(0).Y - l2.GetEndPoint(0).Y);
            }
        }

        public static void Get_intersections(List<Point> Points, ref List<Point> intersections)    //由所有墙轮廓线的端点，获取墙轴线的端点。
        {
            foreach (Point p in Points)
            {
                List<Point> nearbypoints = new List<Point>();
                foreach (Point q in Points)
                {
                    if (P2P_distance(p, q) <= 500 / 304.8)
                    {
                        nearbypoints.Add(q);        //对于每个点p，nearbypoints含有包括了p自身在内所有与p距离在墙厚以内的点。
                    }
                }
                intersections.Add(GetCenter(nearbypoints));                    //得到这些点的中心，可作为墙轴线的端点，也即intersections。
            }
            for (int i = 0; i < intersections.Count; i++)
            {
                for (int j = intersections.Count - 1; j > i; j--)
                {
                    if (Math.Abs(intersections[i].Coord.X - intersections[j].Coord.X) < 0.000328
                        && Math.Abs(intersections[i].Coord.Y - intersections[j].Coord.Y) < 0.000328
                        && Math.Abs(intersections[i].Coord.Z - intersections[j].Coord.Z) < 0.000328)
                    {
                        intersections.RemoveAt(j);
                    }
                }
            }
        }

        public static void CreateWallAxes(List<Point> intersections, ref List<Line> Lines)     //输入墙轴线端点的list，两两连接得到可能的实际墙轴线。
        {
            for (int i = intersections.Count - 1; i > 0; i--)
            {
                Point point_temp = intersections[i];
                intersections.RemoveAt(i);
                foreach (Point intersection_2 in intersections)
                {
                    Lines.Add(Line.CreateBound(point_temp.Coord, intersection_2.Coord));
                }
            }
        }

        public static bool Is_same_direction(XYZ direction1, XYZ direction2)
        {
            return ((IsEqual(direction1.X, direction2.X) && IsEqual(direction1.Y, direction2.Y) && IsEqual(direction1.Z, direction2.Z)) ||
                   (IsEqual(direction1.X, -direction2.X) && IsEqual(direction1.Y, -direction2.Y) && IsEqual(direction1.Z, -direction2.Z)));
        }

        public static bool Is_parallel(Line l1, Line l2)       //判断两直线是否平行。
        {
            if (Is_same_direction(l1.Direction, l2.Direction))
            {
                return true;
            }
            else return false;
        }

        public static bool Isnot_far(Line l1, Line l2)
        {
            double d = L2L_distance(l1, l2);
            if (d < 500 / 304.8 && (!IsEqual(d, 0)))
            {
                return true;
            }
            else return false;
        }

        public static bool Isnot_staggered(Line l1, Line l2)       //判断两边线是否错开。
        {
            Line shortline, longline;
            if (l1.Length < l2.Length || IsEqual(l1.Length, l2.Length))
            {
                shortline = l1;
                longline = l2;
            }
            else
            {
                shortline = l2;
                longline = l1;
            }

            double mincoord, maxcoord;
            if (IsEqual(shortline.Direction.X, 0))
            {
                maxcoord = Math.Max(longline.GetEndPoint(0).Y, longline.GetEndPoint(1).Y);
                mincoord = Math.Min(longline.GetEndPoint(0).Y, longline.GetEndPoint(1).Y);
                if ((shortline.GetEndPoint(0).Y > mincoord - 1e-3 && shortline.GetEndPoint(0).Y < maxcoord + 1e-3) ||
                    (shortline.GetEndPoint(1).Y > mincoord - 1e-3 && shortline.GetEndPoint(1).Y < maxcoord + 1e-3))
                {
                    return true;
                }
                else return false;
            }
            else
            {
                maxcoord = Math.Max(longline.GetEndPoint(0).X, longline.GetEndPoint(1).X);
                mincoord = Math.Min(longline.GetEndPoint(0).X, longline.GetEndPoint(1).X);
                if ((shortline.GetEndPoint(0).X > mincoord - 1e-3 && shortline.GetEndPoint(0).X < maxcoord + 1e-3) ||
                    (shortline.GetEndPoint(1).X > mincoord - 1e-3 && shortline.GetEndPoint(1).X < maxcoord + 1e-3))
                {
                    return true;
                }
                else return false;
            }
        }

        public static void Get_connect_points(WS ws, ref List<Point> connectpoints)       //获取墙段的邻接节点。
        {
            Line BL1 = ws.BL1, BL2 = ws.BL2;

            if (IsEqual(BL1.Direction.X, 0))
            {
                double mincoord, maxcoord;
                maxcoord = Math.Max(BL2.GetEndPoint(0).Y, BL2.GetEndPoint(1).Y);
                mincoord = Math.Min(BL2.GetEndPoint(0).Y, BL2.GetEndPoint(1).Y);
                if (BL1.GetEndPoint(0).Y > mincoord - 1e-3 && BL1.GetEndPoint(0).Y < maxcoord + 1e-3)
                {
                    connectpoints.Add(Point.Create(BL1.GetEndPoint(0)));
                }
                if (BL1.GetEndPoint(1).Y > mincoord - 1e-3 && BL1.GetEndPoint(1).Y < maxcoord + 1e-3)
                {
                    connectpoints.Add(Point.Create(BL1.GetEndPoint(1)));
                }
            }

            if (IsEqual(BL2.Direction.X, 0))
            {
                double mincoord, maxcoord;
                maxcoord = Math.Max(BL1.GetEndPoint(0).Y, BL1.GetEndPoint(1).Y);
                mincoord = Math.Min(BL1.GetEndPoint(0).Y, BL1.GetEndPoint(1).Y);
                if (BL2.GetEndPoint(0).Y > mincoord - 1e-3 && BL2.GetEndPoint(0).Y < maxcoord + 1e-3)
                {
                    connectpoints.Add(Point.Create(BL2.GetEndPoint(0)));
                }
                if (BL2.GetEndPoint(1).Y > mincoord - 1e-3 && BL2.GetEndPoint(1).Y < maxcoord + 1e-3)
                {
                    connectpoints.Add(Point.Create(BL2.GetEndPoint(1)));
                }
            }

            if (IsEqual(BL1.Direction.Y, 0))
            {
                double mincoord, maxcoord;
                maxcoord = Math.Max(BL1.GetEndPoint(0).X, BL1.GetEndPoint(1).X);
                mincoord = Math.Min(BL1.GetEndPoint(0).X, BL1.GetEndPoint(1).X);
                if (BL2.GetEndPoint(0).X > mincoord - 1e-3 && BL2.GetEndPoint(0).X < maxcoord + 1e-3)
                {
                    connectpoints.Add(Point.Create(BL2.GetEndPoint(0)));
                }
                if (BL2.GetEndPoint(1).X > mincoord - 1e-3 && BL2.GetEndPoint(1).X < maxcoord + 1e-3)
                {
                    connectpoints.Add(Point.Create(BL2.GetEndPoint(1)));
                }
            }

            if (IsEqual(BL2.Direction.Y, 0))
            {
                double mincoord, maxcoord;
                maxcoord = Math.Max(BL2.GetEndPoint(0).X, BL2.GetEndPoint(1).X);
                mincoord = Math.Min(BL2.GetEndPoint(0).X, BL2.GetEndPoint(1).X);
                if (BL1.GetEndPoint(0).X > mincoord - 1e-3 && BL1.GetEndPoint(0).X < maxcoord + 1e-3)
                {
                    connectpoints.Add(Point.Create(BL1.GetEndPoint(0)));
                }
                if (BL1.GetEndPoint(1).X > mincoord - 1e-3 && BL1.GetEndPoint(1).X < maxcoord + 1e-3)
                {
                    connectpoints.Add(Point.Create(BL1.GetEndPoint(1)));
                }
            }
        }

        public static bool Is_neighbor(WS ws1, WS ws2)
        {
            List<Point> pointlist1 = new List<Point>();
            List<Point> pointlist2 = new List<Point>();
            Get_connect_points(ws1, ref pointlist1);
            Get_connect_points(ws2, ref pointlist2);
            int i = 0;
            foreach (Point p1 in pointlist1)
            {
                foreach (Point p2 in pointlist2)
                {
                    if (IsSameXYZ(p1.Coord, p2.Coord))
                    {
                        i++;
                    }
                }
            }
            if (i > 0) return true;
            else return false;
        }

        public static void Cal_Wallaxis(WS ws, ref double X, ref double Y)
        {
            Line l1 = ws.BL1;
            Line l2 = ws.BL2;
            if (IsEqual(l1.Direction.X, 0))
            {
                X = (l1.GetEndPoint(0).X + l2.GetEndPoint(0).X) / 2;
            }
            else
            {
                Y = (l1.GetEndPoint(0).Y + l2.GetEndPoint(0).Y) / 2;
            }
        }

        public static int NotExistsIn(Point p, List<Point> plist)
        {
            int i;
            for (i = 0; i < plist.Count; i++)
            {
                if (IsSameXYZ(p.Coord, plist[i].Coord)) break;
            }
            return i;

        }

        public static void SortByCoord(ref List<Point> pointlist)
        {
            for (int i = 0; i < pointlist.Count - 1; i++)
            {
                for (int j = 0; j < pointlist.Count - 1 - i; j++)
                {
                    if ((pointlist[j].Coord.X + pointlist[j].Coord.Y) >
                        (pointlist[j + 1].Coord.X + pointlist[j + 1].Coord.Y))
                    {
                        var temp = pointlist[j + 1];
                        pointlist[j + 1] = pointlist[j];
                        pointlist[j] = temp;
                    }
                }
            }
        }

        public static bool Exists_NearbyLine(Line L, List<Line> lines)     //判断轴线附近是否存在墙线，返回一个Boolean值。 
        {
            int i = 0;
            foreach (Line L_wall in lines)
            {
                if (L2L_distance(L, L_wall) < 300 / 308.4 && Is_parallel(L, L_wall) && Isnot_staggered(L, L_wall))     //判断依据：存在至少一条线段，使得其与轴线的距离小于最大墙宽的一半。
                {
                    i++;
                }
            }
            if (i == 0)
            {
                return false;
            }
            else return true;
        }

        public static void Gen_WSlist(List<Line> Walloutlines, ref List<WS> WSlist)
        {
            foreach (Line line in Walloutlines)
            {
                foreach (Line line1 in Walloutlines)
                {
                    if (Is_parallel(line, line1) &&
                        Isnot_staggered(line, line1) &&
                        Isnot_far(line, line1))
                    {
                        WSlist.Add(new WS(line, line1));
                    }
                }
            }
        }

        public static bool IsSameLine(Line l1, Line l2)
        {
            XYZ l1_endpoint1 = l1.GetEndPoint(0);
            XYZ l1_endpoint2 = l1.GetEndPoint(1);
            XYZ l2_endpoint1 = l2.GetEndPoint(0);
            XYZ l2_endpoint2 = l2.GetEndPoint(1);

            return ((IsSameXYZ(l1_endpoint1, l2_endpoint1) && IsSameXYZ(l1_endpoint2, l2_endpoint2)) ||
                    (IsSameXYZ(l1_endpoint1, l2_endpoint2) && IsSameXYZ(l1_endpoint2, l2_endpoint1)));
        }

        public static bool IsSameWS(WS ws1, WS ws2)
        {
            return (IsSameLine(ws1.BL1, ws2.BL1) && IsSameLine(ws1.BL2, ws2.BL2)) ||
                   (IsSameLine(ws1.BL1, ws2.BL2) && IsSameLine(ws1.BL2, ws2.BL1));
        }
    }

    public class WS
    {
        public Line BL1, BL2;
        public List<int> KnotIndexs;
        public double X = 0;
        public double Y = 0;

        public WS()
        {
            KnotIndexs = new List<int>();
        }

        public WS(Line BL1, Line BL2)
        {
            this.BL1 = BL1;
            this.BL2 = BL2;
            KnotIndexs = new List<int>();
        }
    }
}