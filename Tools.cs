// Variables necesarias en todos los Comandos Externos
UIApplication uiApp = commandData.Application;
UIDocument uiDoc = uiApp.ActiveUIDocument;
Application app = uiApp.Application;
Document doc = uiDoc.Document;

// Obtener la referencia de un objeto seleccionado
Reference r = uiDoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, "Seleccionar un elemento");
// Obtener el elemento elegido
Element e = doc.GetElement(r.ElementId);

// Obtener Parametros
Parameter pComentario = _elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
Parameter pMarca = _elem.get_Parameter(BuiltInParameter.DOOR_NUMBER);

// Iniciar Nueva Transaccion
Transaction t = new Transaction(_doc, "Cambiar Parametro");
t.Start();

// Modificar Parametros
pComentario.Set(comentario);
pMarca.Set(marca);

t.Commit();

// Filtros y Colecciones
// Encontrar todos los ejemplares de Muros del Documento
ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
FilteredElementCollector collector = new FilteredElementCollector(document);
List<Element> lista = collector.WherePasses(filter).WhereElementIsNotElementType().ToList();
// Otras variantes
// Obtener todas las Topografías
FilteredElementCollector collector = new FilteredElementCollector(doc);
List<Element> lstTopo = collector.OfClass(typeof(TopographySurface)).ToList();

// LinQ: se necesita igualmente un Collector
FilteredElementCollector collector = new FilteredElementCollector(doc);
List<Element> elements = collector.WhereElementIsNotElementType().ToList();
List<Element> lstInstance = (from elem in elements
                                         where elem.Category != null
                                         && elem.Category.Id != (new ElementId(BuiltInCategory.OST_Cameras))
                                         && elem.Category.Id != (new ElementId(BuiltInCategory.OST_StackedWalls))
                                         select elem).ToList();

// Seleccionar elemento en la Interfaz
List<ElementId> lista = new List<ElementId>();
lista.Add(elem.Id);
uiDoc.Selection.SetElementIds(lista);

// Obtener todos los Elementos del Modelo
public static List<Element> GetAllElements(Document doc)
{
    FilteredElementCollector collector = new FilteredElementCollector(doc);
    return collector.WhereElementIsNotElementType().ToList();
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
// Obtener una Lista de FamilySymbol
public static List<FamilySymbol> ObtenerListaTiposFamiliaModelo(Document doc)
{
    List<FamilySymbol> lst = new List<FamilySymbol>();
    FilteredElementCollector collector = new FilteredElementCollector(doc);
    List<Element> lstElem = collector.OfClass(typeof(FamilySymbol)).ToList();
    foreach (Element elem in lstElem)
    {
        // Convertir el Elemento en FamilySymbol
        FamilySymbol fm = elem as FamilySymbol;
        // Verificar si es una Familia de Modelo
        if (fm.Category.CategoryType == CategoryType.Model)
        {
            lst.Add(fm);
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

// Obtener una Familia a partir de su nombre
public static FamilySymbol ObtenerTipoFamiliaPorNombre(Document doc, string name)
{
    FamilySymbol family = null;
    foreach (FamilySymbol sym in GetAllFamilySymbol(doc))
    {
        if (sym.Name == name)
        {
            family = sym;
        }
    }
    return family;
}

// Obtener un Nivel a partir de su nombre
public static Level ObtenerNivelPorNombre(Document doc, string name)
{
    Level lvl = null;
    foreach (Level level in GetAllLevels(doc))
    {
        if (level.Name == name)
        {
            lvl = level;
        }
    }
    return lvl;
}

// Crear Ejemplares de Familia. Se debe crear una lista de FamilyInstanceCreationData
// Se debe referenciar <<using Autodesk.Revit.Creation;>>
FamilyInstanceCreationData ficreationdata = new FamilyInstanceCreationData(pointXYZ, familySymbol, 
                        level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
_doc.Create.NewFamilyInstances2(lstData);

// TEMA FM
// Crear una lista de Elementos para Trackear
public static List<FamilyInstance> ObtenerFamiliasParaFm(Document doc)
{
    List<FamilyInstance> lstInstances = new List<FamilyInstance>();
    FilteredElementCollector col = new FilteredElementCollector(doc);
    var familyInstances = col.WhereElementIsNotElementType().WhereElementIsViewIndependent().OfClass(typeof(FamilyInstance));
    List<Element> lst = (from elem in familyInstances
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
    foreach (var elem in lst)
    {
        FamilyInstance fam = elem as FamilyInstance;
        lstInstances.Add(fam);
    }
    return lstInstances;
}

// Crear una lista de Rooms
public static List<Room> ObtenerHabitaciones(Document doc)
{
    List<Room> lstRooms = new List<Room>();
    FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(SpatialElement));

    foreach (SpatialElement e in collector)
    {
        Room room = e as Room;

        if (null != room)
        {
            if (null != room.Level)
            {
                lstRooms.Add(room);
            }
        }
    }
    return lstRooms;
}

// Exportar un TreeView a un archivo CSV
public static void ExportarTreeViewCsv(TreeView tree, string rutaArchivo)
{
    // Crear un stringBuilder
    StringBuilder sb = new StringBuilder();

    // Agregar los titulos
    sb.AppendLine("Habitacion;Familia");
  
    // Recorrer las ramas del TreeView
    foreach (TreeNode nodeRoom in tree.Nodes[0].Nodes)
    {
        foreach (TreeNode nodeFamilia in nodeRoom.Nodes)
        {
            // Crear la linea de texto del archivo CSV
            string linea = nodeRoom.Text + ";" + nodeFamilia.Text;
            // Agregar al stringbuilder
            sb.AppendLine(linea);
        }
    }
    File.WriteAllText(rutaArchivo, sb.ToString());
}

// Obtener una Habitación de un TreeNodo
public static Room ObtenerRoomDesdeTreeNodo(TreeNode nodo, Document doc)
{
    string id = nodo.Name;
    ElementId elemId = new ElementId(Convert.ToInt32(id));
    Element elem = doc.GetElement(elemId);
    return elem as Room;
}

// Obtener un Elemento de un TreeNodo
public static Element ObtenerElementDesdeTreeNodo(TreeNode nodo, Document doc)
{
    string id = nodo.Name;
    ElementId elemId = new ElementId(Convert.ToInt32(id));
    return doc.GetElement(elemId);
}

// Exportar el TreeView Habitaciones a un Excel
public static void ExportarTreeViewExcel(TreeView tree, string rutaArchivo, Document doc)
{
    // Crear una DataTable para guardar los datos
    DataTable dt = new DataTable();

    // Agregar Columnas a la DataTable
    dt.Columns.Add("Habitación N°");
    dt.Columns.Add("Habitación Nombre");
    dt.Columns.Add("Activo Categoria");
    dt.Columns.Add("Activo Nombre");
    dt.Columns.Add("Activo ID");

    // Agregar Filas a la DataTable
    foreach (TreeNode node1 in tree.Nodes[0].Nodes)
    {
        dt.Rows.Add();
        foreach (TreeNode node2 in node1.Nodes)
        {
            dt.Rows.Add();
        }
    }

    // Crear el Libro de Excel
    XLWorkbook wb = new XLWorkbook();

    // Crear la Hoja de Excel con la DataTable
    wb.Worksheets.Add(dt, "Activos");

    // Obtener la Hoja de Excel creada
    var ws = wb.Worksheet(1);
    int row = 2;

    // Recorrer el TreeView y completar las celdas de la Hoja de Excel
    foreach (TreeNode node1 in tree.Nodes[0].Nodes)
    {
        // Nodo de Habitación
        Room habitacion = Tools.ObtenerRoomDesdeTreeNodo(node1, doc);
        ws.Cell(row, 1).Value = habitacion.Number;
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 2).Value = habitacion.Name;
        ws.Cell(row, 2).Style.Font.Bold = true;
        row++;

        // Nodos de Activos
        foreach (TreeNode node2 in node1.Nodes)
        {
            Element activo = Tools.ObtenerElementDesdeTreeNodo(node2, doc);
            ws.Cell(row, 1).Value = habitacion.Number;
            ws.Cell(row, 2).Value = habitacion.Name;
            ws.Cell(row, 3).Value = activo.Category.Name;
            ws.Cell(row, 4).Value = activo.Name;
            ws.Cell(row, 5).Value = activo.Id.IntegerValue.ToString();
            row++;
        }
    }

    // Guardar el libro de excel
    wb.SaveAs(rutaArchivo);
}

// Obtener una Lista de FAces de un elemento
public static List<Face> getFaces(Document doc, Element e)
{
    List<Face> lstFaces = new List<Face>();
    Options opt = new Options();
    opt.DetailLevel = ViewDetailLevel.Fine;
    opt.ComputeReferences = true;
    GeometryElement geom = e.get_Geometry(opt);
    foreach (GeometryObject geobj in geom)
    {
        Solid geomSolid = geobj as Solid;
        if (geomSolid != null)
        {
            foreach (Face geoFace in geomSolid.Faces)
            {
                lstFaces.Add(geoFace);
            }
        }
    }
    return lstFaces;
}

// Importar Puntos desde CSV
public static List<XYZ> ImportarPuntosCSV(string rutaArchivo)
{
    // Crear Lista de Puntos vacia
    List<XYZ> listaPuntos = new List<XYZ>();

    // Leer líneas del archivo CSV
    string[] lines = File.ReadAllLines(rutaArchivo);

    // Quitar la primera línea de titulos
    string[] lines2 = lines.Skip(1).ToArray();

    for (int i = 0; i < lines2.Length; i++)
    {
        string[] values = Regex.Split(lines2[i], ";");
        try
        {
            double x = Convert.ToDouble(values[0]);
            double y = Convert.ToDouble(values[1]);
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
