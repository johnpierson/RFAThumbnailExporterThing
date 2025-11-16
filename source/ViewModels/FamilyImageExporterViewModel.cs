using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CommunityToolkit.Mvvm.Input;

namespace FamilyImageExporter.ViewModels
{
    public sealed class FamilyImageExporterViewModel : ObservableObject
    {
        private readonly UIApplication _uiApplication;
        private string _sourceFolderPath = string.Empty;
        private string _outputFolderPath = string.Empty;
        private string _statusMessage = "Ready";
        private bool _isExporting;
        private int _progressValue;

        public FamilyImageExporterViewModel(UIApplication uiApplication)
        {
            _uiApplication = uiApplication;
            SelectSourceFolderCommand = new RelayCommand(SelectSourceFolder);
            SelectOutputFolderCommand = new RelayCommand(SelectOutputFolder);
            ExportCommand = new RelayCommand(ExportThumbnails, () => !IsExporting && !string.IsNullOrEmpty(SourceFolderPath));
        }

        public string SourceFolderPath
        {
            get => _sourceFolderPath;
            set
            {
                if (SetProperty(ref _sourceFolderPath, value))
                {
                    (ExportCommand as RelayCommand)?.NotifyCanExecuteChanged();
                }
            }
        }

        public string OutputFolderPath
        {
            get => _outputFolderPath;
            set => SetProperty(ref _outputFolderPath, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public bool IsExporting
        {
            get => _isExporting;
            set
            {
                if (SetProperty(ref _isExporting, value))
                {
                    (ExportCommand as RelayCommand)?.NotifyCanExecuteChanged();
                }
            }
        }

        public int ProgressValue
        {
            get => _progressValue;
            set => SetProperty(ref _progressValue, value);
        }

        public ICommand SelectSourceFolderCommand { get; }
        public ICommand SelectOutputFolderCommand { get; }
        public ICommand ExportCommand { get; }

        private void SelectSourceFolder()
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "Select folder containing family files (.rfa)",
                ShowNewFolderButton = false
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                SourceFolderPath = dialog.SelectedPath;
                if (string.IsNullOrEmpty(OutputFolderPath))
                {
                    OutputFolderPath = Path.Combine(SourceFolderPath, "Thumbnails");
                }
            }
        }

        private void SelectOutputFolder()
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "Select output folder for thumbnails",
                ShowNewFolderButton = true
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                OutputFolderPath = dialog.SelectedPath;
            }
        }

        private void ExportThumbnails()
        {
            if (string.IsNullOrEmpty(SourceFolderPath) || !Directory.Exists(SourceFolderPath))
            {
                StatusMessage = "Invalid source folder path";
                return;
            }

            if (string.IsNullOrEmpty(OutputFolderPath))
            {
                OutputFolderPath = Path.Combine(SourceFolderPath, "Thumbnails");
            }

            IsExporting = true;
            StatusMessage = "Starting export...";
            ProgressValue = 0;

            try
            {
                if (!Directory.Exists(OutputFolderPath))
                {
                    Directory.CreateDirectory(OutputFolderPath);
                }

                var familyFiles = Directory.GetFiles(SourceFolderPath, "*.rfa", SearchOption.TopDirectoryOnly);
                int successCount = 0;
                int totalCount = familyFiles.Length;

                StatusMessage = $"Found {totalCount} family files. Starting export...";

                foreach (var familyPath in familyFiles)
                {
                    try
                    {
                        StatusMessage = $"Processing: {Path.GetFileName(familyPath)}";

                        if (ExportFamilyThumbnail(familyPath, OutputFolderPath))
                        {
                            successCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        StatusMessage = $"Error processing {Path.GetFileName(familyPath)}: {ex.Message}";
                    }
                }

                ProgressValue = 100;
                StatusMessage = $"Export complete. {successCount} of {totalCount} families exported successfully.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error: {ex.Message}";
            }
            finally
            {
                IsExporting = false;
            }
        }

        private bool ExportFamilyThumbnail(string familyPath, string outputFolder)
        {
            Document familyDoc = null;
            try
            {
                // Open the family document
                familyDoc = _uiApplication.OpenAndActivateDocument(ModelPathUtils.ConvertUserVisiblePathToModelPath(familyPath), new OpenOptions(), false).Document;

                if (familyDoc == null)
                {
                    return false;
                }

                // Get a 3D view or create one if it doesn't exist
                View3D view3D = GetOrCreate3DView(familyDoc);

                if (view3D == null)
                {
                    return false;
                }

                // Get UIDocument for the family document
                UIDocument uiDoc = _uiApplication.ActiveUIDocument;

                // Make sure the view is active
                if (uiDoc != null)
                {
                    uiDoc.ActiveView = view3D;
                }

                // Ensure all elements are visible in the view
                MakeElementsVisible(view3D);

                // Export the image
                var fileName = Path.GetFileNameWithoutExtension(familyPath);
                var outputPath = Path.Combine(outputFolder, $"{fileName}.png");

                var exportOptions = new ImageExportOptions
                {
                    FilePath = outputPath,
                    ZoomType = ZoomFitType.FitToPage,
                    PixelSize = 1024,
                    ImageResolution = ImageResolution.DPI_150,
                    FitDirection = FitDirectionType.Horizontal,
                    ShadowViewsFileType = ImageFileType.PNG,
                    HLRandWFViewsFileType = ImageFileType.PNG,
                    ExportRange = ExportRange.CurrentView,
                    ViewName = view3D.Name
                };

                familyDoc.ExportImage(exportOptions);

                return File.Exists(outputPath);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to export thumbnail for {Path.GetFileName(familyPath)}: {ex.Message}", ex);
            }
            finally
            {
                // Close the family document without saving
                if (familyDoc != null && familyDoc.IsValidObject)
                {
                    familyDoc.Close(false);
                }
            }
        }

        private View3D GetOrCreate3DView(Document doc)
        {
            // Try to find an existing 3D view
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate && v.CanBePrinted);

            var view3D = collector.FirstOrDefault();

            if (view3D == null)
            {
                // Create a new 3D view if none exists
                using (var trans = new Transaction(doc, "Create 3D View"))
                {
                    trans.Start();

                    // Get the 3D view family type
                    var viewFamilyType = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewFamilyType))
                        .Cast<ViewFamilyType>()
                        .FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);

                    if (viewFamilyType != null)
                    {
                        // Create an isometric 3D view
                        view3D = View3D.CreateIsometric(doc, viewFamilyType.Id);
                    }

                    trans.Commit();
                }
            }

            return view3D;
        }

        private void MakeElementsVisible(View3D view)
        {
            var doc = view.Document;
            using (var trans = new Transaction(doc, "Show All Elements"))
            {
                trans.Start();

                // Get all family instances and make them visible
                var elements = new FilteredElementCollector(doc, view.Id)
                    .WhereElementIsNotElementType()
                    .ToElements();

                foreach (var element in elements)
                {
                    try
                    {
                        if (element.CanBeHidden(view))
                        {
                            view.UnhideElements(new[] { element.Id });
                        }
                    }
                    catch
                    {
                        // Ignore elements that can't be unhidden
                    }
                }

                // Set view display mode to shaded or realistic for better visibility
                try
                {
                    view.get_Parameter(BuiltInParameter.VIEW_MODEL_DISPLAY_MODE)?.Set((int)DisplayStyle.Shading);
                }
                catch
                {
                    // Ignore if parameter can't be set
                }

                trans.Commit();
            }
        }
    }
}