using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Nice3point.Revit.Toolkit.External;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace RevitView2svg.Commands;

[UsedImplicitly]
[Transaction(TransactionMode.Manual)]
public class CmdExportToSvg : ExternalCommand
{
    public override void Execute()
    {
        var uidoc = ExternalCommandData.Application.ActiveUIDocument;
        var document = uidoc.Document;
        var activeView = document.ActiveView;

        if (activeView is not ViewPlan viewPlan || activeView.IsTemplate || !activeView.IsValidObject)
            return;

        List<Line> geometryLines = [];
        Options geomOptions = new()
        {
            ComputeReferences = false,
            IncludeNonVisibleObjects = false,
            View = activeView
        };

        var elements = new FilteredElementCollector(document, activeView.Id).WhereElementIsNotElementType();

        foreach (Element element in elements)
        {
            GeometryElement geomElem = element.get_Geometry(geomOptions);
            if (geomElem == null) continue;
            try
            {
                foreach (GeometryObject geomObj in geomElem)
                    ExtractLinesFromGeometry(geomObj, geometryLines);
            }
            catch
            {
                var e = geomElem;
            }
        }

        if (geometryLines.Count == 0)
        {
            TaskDialog.Show("SVG Export", "No lines found in view.");
            return;
        }

        XYZ origin = activeView.Origin;
        XYZ right = activeView.RightDirection;
        XYZ up = activeView.UpDirection;

        List<XYZ> allPoints = [];
        foreach (Line l in geometryLines)
        {
            if (l == null) continue;
            allPoints.Add(l.GetEndPoint(0));
            allPoints.Add(l.GetEndPoint(1));
        }

        var projected = allPoints.Select(p =>
        {
            double x = right.DotProduct(p - origin);
            double y = up.DotProduct(p - origin);
            return new UV(x, y);
        }).ToList();

        double minX = projected.Min(p => p.U);
        double maxX = projected.Max(p => p.U);
        double minY = projected.Min(p => p.V);
        double maxY = projected.Max(p => p.V);

        double scale = 100.0;
        double width = (maxX - minX) * scale;
        double height = (maxY - minY) * scale;

        StringBuilder svg = new();
        svg.AppendLine($"<svg xmlns='http://www.w3.org/2000/svg' width='{width}' height='{height}' viewBox='0 0 {width} {height}'>");
        svg.AppendLine("<g stroke='black' stroke-width='1' fill='none'>");

        foreach (Line l in geometryLines)
        {
            if (l == null) continue;
            XYZ p1 = l.GetEndPoint(0);
            XYZ p2 = l.GetEndPoint(1);

            double x1 = (right.DotProduct(p1 - origin) - minX) * scale;
            double y1 = (maxY - up.DotProduct(p1 - origin)) * scale;
            double x2 = (right.DotProduct(p2 - origin) - minX) * scale;
            double y2 = (maxY - up.DotProduct(p2 - origin)) * scale;

            svg.AppendLine($"<line x1='{x1}' y1='{y1}' x2='{x2}' y2='{y2}' />");
        }

        svg.AppendLine("</g>");
        svg.AppendLine("</svg>");

        string folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string path = Path.Combine(folder, "active_view_export.svg");
        File.WriteAllText(path, svg.ToString());

        TaskDialog downloadDialog = new("SVG file is ready")
        {
            MainContent = $"Open SVG file",
            AllowCancellation = false
        };
        downloadDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                   "Open SVG file",
                   "This option will open SVG file.");
        downloadDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                   "Navigate to SVG file",
                   "This option will open SVG file directory.");
        downloadDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3,
                   "Close");

        switch (downloadDialog.Show())
        {
            case TaskDialogResult.CommandLink1:
                try
                {
                    Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", $"Failed to open file: {ex.Message}");
                }
                break;

            case TaskDialogResult.CommandLink2:
                Process.Start("explorer.exe", $"/select,\"{path}\"");
                break;

            default:
                break;
        }
    }

    private void ExtractLinesFromGeometry(GeometryObject geometryObject, List<Line> lines)
    {
        if (geometryObject is Line line)
            lines.Add(line);
        else if (geometryObject is Arc arc)
        {
            int segments = 10;
            double start = arc.GetEndParameter(0);
            double end = arc.GetEndParameter(1);

            for (int i = 0; i < segments; i++)
            {
                double t1 = start + (end - start) * i / segments;
                double t2 = start + (end - start) * (i + 1) / segments;

                XYZ pt1 = arc.Evaluate(t1, false);
                XYZ pt2 = arc.Evaluate(t2, false);
                lines.Add(CreateBoundLine(pt1, pt2));
            }
        }
        else if (geometryObject is PolyLine poly)
        {
            IList<XYZ> pts = poly.GetCoordinates();
            for (int i = 0; i < pts.Count - 1; i++)
                lines.Add(CreateBoundLine(pts[i], pts[i + 1]));
        }
        else if (geometryObject is GeometryInstance instance)
        {
            foreach (GeometryObject instObj in instance.GetInstanceGeometry())
                ExtractLinesFromGeometry(instObj, lines);
        }
        else if (geometryObject is Solid solid)
        {
            foreach (Face face in solid.Faces)
                foreach (EdgeArray edgeArray in face.EdgeLoops)
                    foreach (Edge edge in edgeArray)
                    {
                        Curve curve = edge.AsCurve();
                        if (curve.IsBound)
                            if (curve is Line l)
                                lines.Add(l);
                            else
                                TessellateCurve(curve, lines);
                    }
        }
    }

    private void TessellateCurve(Curve curve, List<Line> lines)
    {
        IList<XYZ> points = curve.Tessellate();
        for (int i = 0; i < points.Count - 1; i++)
            lines.Add(CreateBoundLine(points[i], points[i + 1]));
    }

    private Line CreateBoundLine(XYZ pt1, XYZ pt2)
    {
        try
        {
            var line = Line.CreateBound(pt1, pt2);
            return line;
        }
        catch (Exception)
        {
            return null;
        }
    }
}