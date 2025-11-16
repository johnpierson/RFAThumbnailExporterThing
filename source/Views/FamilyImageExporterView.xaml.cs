using FamilyImageExporter.ViewModels;

namespace FamilyImageExporter.Views
{
    public sealed partial class FamilyImageExporterView
    {
        public FamilyImageExporterView(FamilyImageExporterViewModel viewModel)
        {
            DataContext = viewModel;
            InitializeComponent();
        }
    }
}