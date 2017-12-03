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
            // Agregar Panel
            RibbonPanel panel1 = application.CreateRibbonPanel("Test");

            // Agregar Botón
            PushButton boton1 = panel1.AddItem(new PushButtonData("Boton1", "Hola Mundo", ExecutingAssemblyPath,
                "AddinHolaMundo.Class1")) as PushButton;
            PushButton boton2 = panel1.AddItem(new PushButtonData("Boton2", "Seleccionar Elemento", ExecutingAssemblyPath,
                "AddinHolaMundo.Class2")) as PushButton;
            PushButton boton3 = panel1.AddItem(new PushButtonData("Boton3", "Colecciones", ExecutingAssemblyPath,
                "AddinHolaMundo.Class3")) as PushButton;

            // Agregar la imagen al botón
            boton1.LargeImage = new BitmapImage(new Uri("pack://application:,,,/AddinHolaMundoTest;component/Resources/Imagen32.png"));
            boton1.ToolTip = "Hola mundo";
            boton1.LongDescription = "Esta addin lanza un formulario de Hola Mundo";
            boton2.LargeImage = new BitmapImage(new Uri("pack://application:,,,/AddinHolaMundoTest;component/Resources/Imagen32.png"));
            boton2.ToolTip = "Muestra informacion de un Elemento";
            boton2.LongDescription = "Este comando muestra un formulario con los Paráetros del elemento seleccionado";
            boton3.LargeImage = new BitmapImage(new Uri("pack://application:,,,/AddinHolaMundoTest;component/Resources/Imagen32.png"));
            boton3.ToolTip = "Muestra Colecciones del Proyecto";
            boton2.LongDescription = "Este comando muestra un formulario con distintas colecciones del Proyecto";

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
