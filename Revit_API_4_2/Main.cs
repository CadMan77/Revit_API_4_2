//Для всех труб выведите в Excel-таблицу следующие значения:
//имя типа, наружный диаметр, внутренний диаметр, длина.

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
using System.Windows.Forms;

namespace Revit_API_4_2
{
    [Transaction(TransactionMode.Manual)]

    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileName = "PipeInfo_(Revit-API_Lab4-2).XLSX";

            try
            {

            var pipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .ToList();

            string excelPath = Path.Combine(desktopPath, fileName);

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Трубы");

                int rowIndex = 0;

                //"Имя типа;Наружный диаметр, мм;Внутренний диаметр, мм;Длина, мм"

                foreach (var pipe in pipes)
                {

                    //Parameter type = elem.LookupParameter("Type");
                    //Parameter outDia = elem.LookupParameter("Outside Diameter");
                    //Parameter inDia = elem.LookupParameter("Inside Diameter");
                    //Parameter length = elem.LookupParameter("Length");

                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.PipeType.Name);
                    double outDiaFeet = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble();
                    sheet.SetCellValue(rowIndex, columnIndex: 1, UnitUtils.ConvertFromInternalUnits(outDiaFeet, UnitTypeId.Millimeters));
                    double inDiaFeet = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble();
                    sheet.SetCellValue(rowIndex, columnIndex: 2, UnitUtils.ConvertFromInternalUnits(inDiaFeet, UnitTypeId.Millimeters));
                    double pipeLength = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                    sheet.SetCellValue(rowIndex, columnIndex: 3, UnitUtils.ConvertFromInternalUnits(pipeLength, UnitTypeId.Millimeters));
                    rowIndex++;
                }

                workbook.Write(stream);
                workbook.Close();
            }

            System.Diagnostics.Process.Start(excelPath);

            TaskDialog.Show("Выполнено", $"Данные о всех ({pipes.Count}шт.) трубах проекта {doc.Title}.RVT выгружены в файл {fileName}{Environment.NewLine}(см. Рабочий стол)");

            }
            catch
            {
                TaskDialog.Show("Ошибка", "При выполнении команды возник сбой");
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }
}