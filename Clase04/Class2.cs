using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;

namespace AddinHolaMundoTest
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Class2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Variables necesarias
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Application app = uiApp.Application;
            Document doc = uiDoc.Document;

            // Seleccionar un elemento en la pantalla
            TaskDialog.Show("Revit", "Seleccione un elemento");
            // Obtener la referencia seleccionada
            Reference r = uiDoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element,
                "Seleccionar un elemento");
            // Obtener el elemento elegido
            Element e = doc.GetElement(r.ElementId);

            // Crear el objeto formulario
            frmSeleccionarElemento inicio = new frmSeleccionarElemento(e);

            // Mostrar el formulario
            inicio.Show();

            return Result.Succeeded;
        }
    }
}
