using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace AddinHolaMundoTest
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]

    public class App : IExternalApplication
    {
        // Obtener la ruta de la Addin actual
        public static string ExecutingAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;

        public Result OnStartup(UIControlledApplication application)
        {
            // Agregar Panel
            RibbonPanel panel1 = application.CreateRibbonPanel("TEST");

            // Agregar Botón
            PushButton boton1 = panel1.AddItem(new PushButtonData("Boton01", "Hola Mundo", ExecutingAssemblyPath,
                "AddinHolaMundoTest.Class1")) as PushButton;
            PushButton boton2 = panel1.AddItem(new PushButtonData("Boton02", "Seleccionar", ExecutingAssemblyPath,
                "AddinHolaMundoTest.Class2")) as PushButton;

            // Agregar la imagen al botón
            boton1.LargeImage = new BitmapImage(new Uri("pack://application:,,,/AddinHolaMundo;component/Resources/Imagen32.png"));
            boton1.ToolTip = "Hola mundo";
            boton1.LongDescription = "Esta addin lanza un formulario de Hola Mundo";
            boton2.LargeImage = new BitmapImage(new Uri("pack://application:,,,/AddinHolaMundo;component/Resources/Imagen32.png"));
            boton2.ToolTip = "Seleccionar un elemento del modelo";
            boton2.LongDescription = "Este comando muestra información de un elemento del modelo";

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
