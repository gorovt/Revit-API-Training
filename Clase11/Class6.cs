using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.Architecture;

namespace AddinHolaMundo
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Class6 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Variables necesarias en todos los Comandos Externos
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Application app = uiApp.Application;
            Document doc = uiDoc.Document;

            // Obtener listas
            List<Room> lstRooms = Tools.ObtenerHabitaciones(doc);
            List<FamilyInstance> lstFamilies = Tools.ObtenerFamiliasParaFm(doc);

            // Crear el objeto formulario
            frmFm formulario = new frmFm(doc, lstRooms, lstFamilies);
            

            // Mostrar el formulario
            formulario.ShowDialog();

            return Result.Succeeded;
        }
    }
}
