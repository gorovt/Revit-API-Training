using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System.Windows.Forms;
using System.IO;
using ClosedXML.Excel;
using System.Data;

namespace AddinHolaMundo
{
    public class Tools // Clase que nunca se instancia
    {
        // Crear una Lista de Todas las Familias del Proyecto (FamilySymbol)
        public static List<FamilySymbol> ObtenerListaTiposFamiliaModelo(Document doc)
        {
            // Crear una lista de elementos vacia
            List<FamilySymbol> lst = new List<FamilySymbol>();
            
            // Crear un collector de FamilySymbols
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<Element> elementos = collector.OfClass(typeof(FamilySymbol)).ToList();

            // Llenar la Lista de FamilySymbol
            foreach (Element item in elementos)
            {
                if (item.Category.CategoryType == CategoryType.Model)
                {
                    // Convertir el elemento en FamilySymbol
                    FamilySymbol simbolo = item as FamilySymbol;
                    // Rellena la lista
                    lst.Add(simbolo);
                }
            }
            return lst;
        }

        // Crear una lista de Categorias del Modelo
        public static List<Category> ObtenerListaCategoriasModelo(Document doc)
        {
            // Crear una lista vacia
            List<Category> categorias = new List<Category>();

            // Obtener una lista de tipos de familia cargados
            List<FamilySymbol> tipos = ObtenerListaTiposFamiliaModelo(doc);

            // Rellenar la Lista de Category
            foreach (var item in tipos)
            {
                // Verificar que la Categoria NO exista en la Lista
                if (!categorias.Exists(x => x.Name == item.Category.Name))
                {
                    categorias.Add(item.Category);
                }
            }
            return categorias;
        }

        // Crear una lista de Niveles del Proyecto
        public static List<Level> ObtenerListaNiveles(Document doc)
        {
            // Crear una lista vacia de niveles
            List<Level> niveles = new List<Level>();

            // Crear un collector de Niveles
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<Element> elementos = collector.OfClass(typeof(Level)).ToList();

            // Llenar la lista de Niveles
            foreach (var item in elementos)
            {
                // Convierto el Elemento en un Nivel
                Level nivel = item as Level;

                // Lleno la lista de niveles
                niveles.Add(nivel);
            }
            return niveles;
        }

        // Obtener un FamilySymbol según su nombre
        public static FamilySymbol ObtenerFamilySymbolPorNombre(Document doc, string nombreTipo)
        {
            // Creo un FamilySymbol vacio
            FamilySymbol familia = null;

            // Recorrer toda la lista de FamilySymbol buscando el nombre
            foreach (var item in ObtenerListaTiposFamiliaModelo(doc))
            {
                // Verificar el nombre que me dieron
                if (item.Name == nombreTipo)
                {
                    familia = item;
                }
            }
            return familia;
        }

        // Obtener un Nivel según su nombre
        public static Level ObtenerNivelPorNombre(Document doc, string nombreNivel)
        {
            // Crear un nivel vacio
            Level nivel = null;

            // Recoorrer la lista de niveles
            foreach (Level item in ObtenerListaNiveles(doc))
            {
                // Verificar el nombre que me dieron
                if (item.Name == nombreNivel)
                {
                    nivel = item;
                }
            }
            return nivel;
        }

        // Obtener una lista de Habitaciones
        public static List<Room> ObtenerListaHabitaciones(Document doc)
        {
            // Crear una Lista vacia de Rooms
            List<Room> lst = new List<Room>();

            // Creo un Colector de Rooms
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<Element> elementos = collector.OfClass(typeof(SpatialElement)).ToList();

            foreach (Element elem in elementos)
            {
                if (elem as Room != null)
                {
                    // Obtengo una Room
                    Room habitacion = elem as Room;
                    lst.Add(habitacion);
                }
            }

            return lst;
        }

        // Obtener la Lista de Activos
        public static List<FamilyInstance> ObtenerListaActivos(Document doc)
        {
            // Crear lista de Family Instances
            List<FamilyInstance> lst = new List<FamilyInstance>();

            // Collector de FamilyInstances
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<Element> elementos = collector.WhereElementIsNotElementType().WhereElementIsViewIndependent().OfClass(typeof(FamilyInstance)).ToList();

            // Filtrar las categorias que me interesan
            List<Element> lista = (from elem in elementos
                                   where elem.Category.Id == new ElementId(BuiltInCategory.OST_Doors)
                                    || elem.Category.Id == new ElementId(BuiltInCategory.OST_Windows)
                                    || elem.Category.Id == new ElementId(BuiltInCategory.OST_Furniture)
                                    || elem.Category.Id == new ElementId(BuiltInCategory.OST_DuctTerminal)
                                    || elem.Category.Id == new ElementId(BuiltInCategory.OST_PlumbingFixtures)
                                    || elem.Category.Id == new ElementId(BuiltInCategory.OST_MechanicalEquipment)
                                    || elem.Category.Id == new ElementId(BuiltInCategory.OST_ElectricalEquipment)
                                    || elem.Category.Id == new ElementId(BuiltInCategory.OST_LightingDevices)
                                    || elem.Category.Id == new ElementId(BuiltInCategory.OST_LightingFixtures)
                                    || elem.Category.Id == new ElementId(BuiltInCategory.OST_ElectricalFixtures)
                                    || elem.Category.Id == new ElementId(BuiltInCategory.OST_Sprinklers)
                                   select elem).ToList();

            foreach (Element elem in lista)
            {
                FamilyInstance familia = elem as FamilyInstance;
                lst.Add(familia);
            }

            return lst;
        }

        // Exportar TreeView a archivo CSV
        public static void ExportarTreeViewCsv(TreeView tree, string rutaArchivo)
        {
            // Crear un stringBuilder
            StringBuilder sb = new StringBuilder();

            // Agregar titulos al archivo
            sb.AppendLine("Habitacion;Activo");

            // Recorrer las ramas del TreeView y agregar lineas de texto
            foreach (TreeNode nodo1 in tree.Nodes[0].Nodes)
            {
                // Estoy en la rama de las Habitaciones
                foreach (TreeNode nodo2 in nodo1.Nodes)
                {
                    // Estoy en la rama de los activos
                    string linea = nodo1.Text + ";" + nodo2.Text;
                    sb.AppendLine(linea);
                }
            }
            File.WriteAllText(rutaArchivo, sb.ToString());
        }

        // Obtener una Habitación a partir de un TreeNodo
        public static Room ObtenerRoomDesdeTreeNodo(TreeNode nodo, Document doc)
        {
            string nombre = nodo.Name;
            int id = Convert.ToInt32(nombre);
            ElementId elemId = new ElementId(id);
            // Obtener el elemento del modelo
            Element elem = doc.GetElement(elemId);
            Room habitacion = elem as Room;
            return habitacion;
        }

        // Obtener un Activo (Elemento) a partir de un TreeNodo
        public static Element ObtenerElemenoDesdeTreeNode(TreeNode nodo, Document doc)
        {
            string nombre = nodo.Name;
            int id = Convert.ToInt32(nombre);
            ElementId elemId = new ElementId(id);
            // Obtener el elemento del modelo
            Element elem = doc.GetElement(elemId);
            return elem;
        }

        // Exportar el TreeView Activos a Excel
        public static void ExportarTreeViewExcel(TreeView tree, string rutaArchivo, Document doc)
        {
            // Crear una DataTable para guardar los datos
            DataTable dt = new DataTable();

            // Agregar las columnas a la DataTable
            dt.Columns.Add("Habitacion N°");
            dt.Columns.Add("Habitacion Nombre");
            dt.Columns.Add("Activo Categoria");
            dt.Columns.Add("Activo Nombre");
            dt.Columns.Add("Activo ID");

            // Agregar Filas VACIAS a la DataTable
            foreach (TreeNode nodo1 in tree.Nodes[0].Nodes)
            {
                // Estoy en la rama Habitacion
                dt.Rows.Add();
                foreach (TreeNode nodo2 in nodo1.Nodes)
                {
                    // Estoy en el Activo
                    dt.Rows.Add();
                }
            }

            // Crear el Libro de Excel
            XLWorkbook wb = new XLWorkbook();

            // Crear la Hoja de Excel con la DataTable
            wb.Worksheets.Add(dt, "Activos");

            // Obtener la Hoja de Excel recién creada
            var ws = wb.Worksheet(1);
            int row = 2; // Posicion de la fila a modificar

            // Recorrer el TreeView y llenar la Hoja de Excel
            foreach (TreeNode node1 in tree.Nodes[0].Nodes)
            {
                // Estoy en la rama Habitación
                Room habitacion = Tools.ObtenerRoomDesdeTreeNodo(node1, doc);
                // Ecribir la fila de excel
                ws.Cell(row, 1).Value = habitacion.Number;
                ws.Cell(row, 1).Style.Font.Bold = true;
                ws.Cell(row, 2).Value = habitacion.Name;
                ws.Cell(row, 2).Style.Font.Bold = true;
                row++;

                // Recorrer los nodos de Activos
                foreach (TreeNode nodo2 in node1.Nodes)
                {
                    // Estoy en la rama de Activo
                    Element activo = Tools.ObtenerElemenoDesdeTreeNode(nodo2, doc);
                    ws.Cell(row, 1).Value = habitacion.Number;
                    ws.Cell(row, 2).Value = habitacion.Name;
                    ws.Cell(row, 3).Value = activo.Category.Name;
                    ws.Cell(row, 4).Value = activo.Name;
                    ws.Cell(row, 5).Value = activo.Id.IntegerValue.ToString();
                    row++;
                }
            }

            // Guardo el archivo de Excel
            wb.SaveAs(rutaArchivo);
        }

        // Importar Puntos desde Excel
        public static List<XYZ> ImportarPuntosExcel(string rutaArchivo)
        {
            // Crear Lista de Puntos vacia
            List<XYZ> listaPuntos = new List<XYZ>();

            // Libro de Excel y Hojas
            var libro = new XLWorkbook(rutaArchivo);
            var hoja = libro.Worksheet(1);

            // Recorremos las Filas de la Hoja
            foreach (var row in hoja.Rows())
            {
                try
                {
                    double x = Convert.ToDouble(row.Cell(1).Value.ToString());
                    double y = Convert.ToDouble(row.Cell(2).Value.ToString());
                    XYZ punto = new XYZ(x, y, 0);
                    listaPuntos.Add(punto);
                }
                catch (Exception)
                {
                    // No hacer nada
                }
            }

            return listaPuntos;
        }
    }
}
