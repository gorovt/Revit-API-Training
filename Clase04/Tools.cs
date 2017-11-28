// Variables necesarias en todos los Comandos Externos
UIApplication uiApp = commandData.Application;
UIDocument uiDoc = uiApp.ActiveUIDocument;
Application app = uiApp.Application;
Document doc = uiDoc.Document;

// Obtener la referencia de un objeto seleccionado
Reference r = uiDoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, "Seleccionar un elemento");
// Obtener el elemento elegido
Element e = doc.GetElement(r.ElementId);
