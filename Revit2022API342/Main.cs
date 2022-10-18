using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Revit2022API342
{
    [Transaction(TransactionMode.Manual)]

    public class Main : IExternalCommand

    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document document = uidoc.Document;

            List<Pipe> pipeList = new FilteredElementCollector(document)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .ToList();

            List<PipeData> pipesDataList = new List<PipeData>();

            int iD = 0;

            foreach (var pipe in pipeList)
            {
                pipesDataList.Add(new PipeData
                {
                    ID = iD,
                    Name = pipe.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsString(),

                    InnerDiameter = UnitUtils.ConvertFromInternalUnits(
                        pipe.get_Parameter(
                        BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble(),
                        UnitTypeId.Millimeters),

                    OuterDiameter = UnitUtils.ConvertFromInternalUnits(
                        pipe.get_Parameter(
                        BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble(),
                        UnitTypeId.Millimeters),

                    Length = UnitUtils.ConvertFromInternalUnits(
                        pipe.get_Parameter(
                        BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(),
                        UnitTypeId.Millimeters),
                });

                iD++;
            }

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipeInfo.xlsx");

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Лист1");

                int rowIndex = 0;
                foreach (var pipe in pipesDataList)
                {
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.ID);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, pipe.Name);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, pipe.InnerDiameter);
                    sheet.SetCellValue(rowIndex, columnIndex: 3, pipe.OuterDiameter);
                    sheet.SetCellValue(rowIndex, columnIndex: 4, pipe.Length);
                    rowIndex++;
                }
                workbook.Write(stream);
                workbook.Close();
            }

            return Result.Succeeded;
        }
    }
}
