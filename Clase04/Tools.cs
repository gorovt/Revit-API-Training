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
