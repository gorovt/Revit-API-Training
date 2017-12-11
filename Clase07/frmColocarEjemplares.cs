using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB.Structure;

namespace AddinHolaMundo
{
    public partial class frmColocarEjemplares : System.Windows.Forms.Form
    {
        public Autodesk.Revit.DB.Document _doc;

        public frmColocarEjemplares(Autodesk.Revit.DB.Document doc)
        {
            InitializeComponent();
            _doc = doc;
            RellenarCategorias(doc);
            RellenarNiveles(doc);
        }

        private void RellenarCategorias(Autodesk.Revit.DB.Document doc)
        {
            List<Category> lst = Tools.GetModelCategories(doc);
            lst = lst.OrderBy(x => x.Name).ToList();
            // Rellenar el Combo
            foreach (Category cat in lst)
            {
                this.cmbCategorias.Items.Add(cat.Name);
            }
        }

        private void RellenarFamilias(Autodesk.Revit.DB.Document doc, string categoryName)
        {
            this.cmbFamilias.Items.Clear();
            List<FamilySymbol> lst = Tools.GetAllFamilySymbol(doc);
            lst = lst.OrderBy(x => x.Name).ToList();
            // Rellenar el Combo
            foreach (FamilySymbol fm in lst)
            {
                if (fm.Category.Name == categoryName)
                {
                    this.cmbFamilias.Items.Add(fm.Name);
                }
            }
        }

        private void RellenarNiveles(Autodesk.Revit.DB.Document doc)
        {
            List<Level> lst = Tools.GetAllLevels(doc);
            lst = lst.OrderBy(x => x.Name).ToList();
            // Rellenar el Combo
            foreach (Level lvl in lst)
            {
                this.cmbNiveles.Items.Add(lvl.Name);
            }
        }

        private List<XYZ> ObtenerListaPuntos()
        {
            List<XYZ> lst = new List<XYZ>();
            foreach (DataGridViewRow row in this.dgvPuntos.Rows)
            {
                if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    string celda1 = row.Cells[0].Value.ToString();
                    double x = Convert.ToDouble(celda1);
                    x = x / 0.3048;
                    string celda2 = row.Cells[1].Value.ToString();
                    double y = Convert.ToDouble(celda2);
                    y = y / 0.3048;
                    XYZ punto = new XYZ(x, y, 0);
                    lst.Add(punto);
                }
            }
            return lst;
        }

        private void cmbCategorias_SelectedIndexChanged(object sender, EventArgs e)
        {
            string categoryName = this.cmbCategorias.SelectedItem.ToString();
            this.cmbFamilias.Enabled = true;
            RellenarFamilias(_doc, categoryName);
        }

        private void btnColocar_Click(object sender, EventArgs e)
        {
            // Verificar selecciones
            if (this.cmbCategorias.SelectedItem == null)
            {
                TaskDialog.Show("Revit", "Debe seleccionar una Categoría");
                return;
            }
            if (this.cmbFamilias.SelectedItem == null)
            {
                TaskDialog.Show("Revit", "Debe seleccionar una Familia");
                return;
            }
            if (this.cmbNiveles.SelectedItem == null)
            {
                TaskDialog.Show("Revit", "Debe seleccionar un Nivel");
                return;
            }
            if (this.dgvPuntos.Rows.Count <= 1)
            {
                TaskDialog.Show("Revit", "Debe ingresar coordenadas en la Lista de Puntos");
                return;
            }
            // Todo OK, proceder
            // Obtener Lista de Puntos
            List<XYZ> puntos = ObtenerListaPuntos();

            // Obtener elementos seleccionados
            FamilySymbol family = Tools.GetFamilySymbolByName(_doc, this.cmbFamilias.SelectedItem.ToString());
            Level lvl = Tools.GetLevelByName(_doc, this.cmbNiveles.SelectedItem.ToString());
            List<FamilyInstanceCreationData> lstData = new List<FamilyInstanceCreationData>();

            // Crear Transaccion
            Transaction t = new Transaction(_doc, "Crear ejemplares");
            t.Start();
            // Procesar los Puntos y crear una Lista de Data
            foreach (XYZ punto in puntos)
            {
                // Crear objeto FamilyInstanceCreationData
                FamilyInstanceCreationData fiData = new FamilyInstanceCreationData(punto, family, StructuralType.NonStructural);
                lstData.Add(fiData);
            }
            // Colocar ejemplares
            _doc.Create.NewFamilyInstances2(lstData);
            t.Commit();
            TaskDialog.Show("Revit", "Elementos colocados correctamente");
            this.Close();
        }
    }
}
