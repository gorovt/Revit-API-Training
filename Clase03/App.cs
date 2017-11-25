using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace AddinHolaMundo
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class App : IExternalApplication
    {
        // Obtener la ruta de la Addin actual
        public static string ExecutingAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;

        public Result OnStartup(UIControlledApplication application)
        {
            // Agregar el Panel
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("Goro");
            // Agregar un botón
            PushButton pushButton = ribbonPanel.AddItem(new PushButtonData("Boton1",
            "Goro", ExecutingAssemblyPath, "AddinHolaMundo.Class1")) as PushButton;
            // Establecer la imagen del Botón
            pushButton.LargeImage = new BitmapImage(
                new Uri("pack://application:,,,/AddinHolaMundo;component/Resources/Imagen32.png"));
            pushButton.ToolTip = "Hola Mundo";
            pushButton.LongDescription = "Se lanza un formulario Hola Mundo";
            return Result.Succeeded;
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
