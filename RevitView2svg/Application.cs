using Nice3point.Revit.Toolkit.External;
using RevitView2svg.Commands;

namespace RevitView2svg;

[UsedImplicitly]
public class Application : ExternalApplication
{
    public override void OnStartup()
    {
        CreateRibbon();
    }

    private void CreateRibbon()
    {
        var panel = Application.CreatePanel("Export", "SHSS Tools");

        panel.AddPushButton<CmdExportToSvg>("To SVG")
            .SetImage("/RevitView2svg;component/Resources/Icons/ToSVG16.png")
            .SetLargeImage("/RevitView2svg;component/Resources/Icons/ToSVG32.png");
    }
}