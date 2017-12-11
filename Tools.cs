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

// Obtener una Lista de Categorías de Elementos del Modelo
public static List<Category> GetCategories(Document doc)
{
    List<Category> lst = new List<Category>();
    foreach (Element elem in GetAllElements(doc))
    {
        // Verificar que sea una Familia de Modelo
        if (elem.Category.CategoryType == CategoryType.Model)
        {
            // Verificar que no exista en la Lista
            if (!lst.Exists(x => x.Id == elem.Category.Id))
            {
                lst.Add(elem.Category);
            }
        }
    }
    return lst;
}

// Obtener una Lista de Niveles
public static List<Level> GetAllLevels(Document doc)
{
    List<Level> lst = new List<Level>();
    FilteredElementCollector collector = new FilteredElementCollector(doc);
    List<Element> lstElem = collector.OfClass(typeof(Level)).ToList();
    foreach (Element elem in lstElem)
    {
        // Convertir el Elemento en Nivel
        Level lvl = elem as Level;
        lst.Add(lvl);
    }
    return lst;
}

// Obtener una Lista de FamilySymbol
public static List<FamilySymbol> GetAllFamilySymbol(Document doc)
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
            // Verificar si NO existe en la Lista
            if (!lst.Exists(x => x.Id == fm.Id))
            {
                lst.Add(fm);
            }
        }
    }
    return lst;
}

// Obtener una Familia a partir de su nombre
public static FamilySymbol GetFamilySymbolByName(Document doc, string name)
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
public static Level GetLevelByName(Document doc, string name)
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
