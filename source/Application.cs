using FamilyImageExporter.Commands;
using Nice3point.Revit.Toolkit.External;

namespace FamilyImageExporter
{
    /// <summary>
    ///     Application entry point
    /// </summary>
    [UsedImplicitly]
    public class Application : ExternalApplication
    {
        public override void OnStartup()
        {
            CreateRibbon();
        }

        private void CreateRibbon()
        {
            var panel = Application.CreatePanel("Commands", "FamilyImageExporter");

            panel.AddPushButton<StartupCommand>("Execute")
                .SetImage("/FamilyImageExporter;component/Resources/Icons/RibbonIcon16.png")
                .SetLargeImage("/FamilyImageExporter;component/Resources/Icons/RibbonIcon32.png");
        }
    }
}