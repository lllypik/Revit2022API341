using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Revit2022API341
{
    [Transaction(TransactionMode.Manual)]

    public class Main : IExternalCommand

    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document document = uidoc.Document;

            List<Wall> wallsList = new FilteredElementCollector(document)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();

            string wallInfo = string.Empty;

            int wallID = 1;
            foreach (var wall in wallsList)
            {
                string wallIDString = wallID.ToString();
                wallID++;
                string wallType = wall.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString();
                double wallVolume = wall.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble();
                string wallVolumeString = wallVolume.ToString();
                wallInfo += $"{wallIDString}\t{wallType}\t{wallVolumeString}{Environment.NewLine}";
            }

            var saveDialog = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All files (*.*)|*.*",
                FileName = "wallsInfo.csv",
                DefaultExt = ".csv"
            };

            string selectedFilePath = string.Empty;
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = saveDialog.FileName;
            }

            if (string.IsNullOrEmpty(selectedFilePath))
            {
                return Result.Cancelled;
            }

            File.WriteAllText(selectedFilePath, wallInfo);

            return Result.Succeeded;

        }
    }
}
