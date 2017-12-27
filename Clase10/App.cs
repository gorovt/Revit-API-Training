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
            RibbonPanel panel1 = application.CreateRibbonPanel("Revit API");

            // Agregar Botón
            PushButton boton1 = panel1.AddItem(new PushButtonData("Boton1", "Hola Mundo", ExecutingAssemblyPath,
                "AddinHolaMundo.Class1")) as PushButton;
            PushButton boton2 = panel1.AddItem(new PushButtonData("Boton2", "Seleccionar Elemento", ExecutingAssemblyPath,
                "AddinHolaMundo.Class2")) as PushButton;
            PushButton boton3 = panel1.AddItem(new PushButtonData("Boton3", "Modificar parametros", ExecutingAssemblyPath,
                "AddinHolaMundo.Class3")) as PushButton;
            PushButton boton4 = panel1.AddItem(new PushButtonData("Boton4", "Colecciones", ExecutingAssemblyPath,
                "AddinHolaMundo.Class4")) as PushButton;
            PushButton boton5 = panel1.AddItem(new PushButtonData("Boton5", "Colocar Ejemplares", ExecutingAssemblyPath,
                "AddinHolaMundo.Class5")) as PushButton;
            PushButton boton6 = panel1.AddItem(new PushButtonData("Boton6", "FM", ExecutingAssemblyPath,
                "AddinHolaMundo.Class6")) as PushButton;

            // Agregar la imagen al botón
            boton1.LargeImage = new BitmapImage(new Uri("pack://application:,,,/AddinHolaMundo;component/Resources/Imagen32.png"));
            boton1.ToolTip = "Hola mundo";
            boton1.LongDescription = "Esta addin lanza un formulario de Hola Mundo";
            boton2.LargeImage = new BitmapImage(new Uri("pack://application:,,,/AddinHolaMundo;component/Resources/Imagen32.png"));
            boton2.ToolTip = "Muestra informacion de un Elemento";
            boton2.LongDescription = "Este comando muestra un formulario con los Paráetros del elemento seleccionado";
            boton3.LargeImage = new BitmapImage(new Uri("pack://application:,,,/AddinHolaMundo;component/Resources/Imagen32.png"));
            boton3.ToolTip = "Modificar parámetros";
            boton3.LongDescription = "Este comando muestra un formulario que permite modificar el valor de algunos parámetros";
            boton4.LargeImage = new BitmapImage(new Uri("pack://application:,,,/AddinHolaMundo;component/Resources/application-sidebar.png"));
            boton4.ToolTip = "Utilizar Colecciones";
            boton4.LongDescription = "Este comando muestra un formulario que permite visualizar colecciones de elementos";
            boton5.LargeImage = new BitmapImage(new Uri("pack://application:,,,/AddinHolaMundo;component/Resources/table.png"));
            boton5.ToolTip = "Colocar Ejemplares programaticamente";
            boton5.LongDescription = "Este comando muestra un formulario que permite colocar ejemplares en coordenadas XY";
            boton6.LargeImage = new BitmapImage(new Uri("pack://application:,,,/AddinHolaMundo;component/Resources/database.png"));
            boton6.ToolTip = "Elementos en Habitaciones";
            boton6.LongDescription = "Este comando muestra un formulario que permite ver los elementos de cada habitación";

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
