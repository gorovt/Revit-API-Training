using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using System.Linq;

namespace AddinHolaMundo
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Class4 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Variables necesarias en todos los Comandos Externos
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Application app = uiApp.Application;
            Document doc = uiDoc.Document;

            // Crear una Lista de Ejemplares de Muros del Proyecto
            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<Element> lista = collector.WherePasses(filter).WhereElementIsNotElementType().ToList();

            // Crear una Lista de Mobiliarios usando LinQ
            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            List<Element> listaMobiliarios = collector2.WhereElementIsNotElementType().ToList();
            List<Element> listaMuebles = (from elem in listaMobiliarios
                                          where elem.Category != null &&
                                          elem.Category.Id == new ElementId(BuiltInCategory.OST_Furniture)
                                          select elem).ToList();

            // Crear el objeto formulario
            frmColecciones seleccion = new frmColecciones(doc, listaMuebles);

            // Mostrar el formulario
            seleccion.ShowDialog();

            return Result.Succeeded;
        }
    }
}
