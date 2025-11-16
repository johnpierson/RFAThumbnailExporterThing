using Autodesk.Revit.Attributes;
using FamilyImageExporter.ViewModels;
using FamilyImageExporter.Views;
using Nice3point.Revit.Toolkit.External;

namespace FamilyImageExporter.Commands
{
    /// <summary>
    ///     External command entry point
    /// </summary>
    [UsedImplicitly]
    [Transaction(TransactionMode.Manual)]
    public class StartupCommand : ExternalCommand
    {
        public override void Execute()
        {
            var viewModel = new FamilyImageExporterViewModel(UiApplication);
            var view = new FamilyImageExporterView(viewModel);
            view.ShowDialog();
        }
    }
}