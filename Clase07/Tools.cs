using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace AddinHolaMundo
{
    public class Tools
    {
        // Obtener todos los Elementos del Modelo
        public static List<Element> GetAllElements(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            return collector.WhereElementIsNotElementType().ToList();
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

        // Obtener una Lista de Categorías del Modelo
        public static List<Category> GetModelCategories(Document doc)
        {
            List<Category> lst = new List<Category>();
            foreach (FamilySymbol sym in GetAllFamilySymbol(doc))
            {
                // Verificar que tenga Categoría y sea de Modelo
                if (sym.Category != null && sym.Category.CategoryType == CategoryType.Model)
                {
                    // Verificar que no exista en la Lista
                    if (!lst.Exists(x => x.Id == sym.Category.Id))
                    {
                        lst.Add(sym.Category);
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
    }
}
