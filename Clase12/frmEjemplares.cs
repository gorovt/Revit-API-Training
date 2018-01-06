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
    public partial class frmEjemplares : System.Windows.Forms.Form
    {
        public Autodesk.Revit.DB.Document _doc;

        public frmEjemplares(Autodesk.Revit.DB.Document doc)
        {
            InitializeComponent();
            _doc = doc;

            // Rellenar Combo Categorias
            RellenarComboCategorias(doc);
            // Relleno Combo Niveles
            RellenarComboNiveles(doc);
        }

        // Rellenar Combo Categorias
        private void RellenarComboCategorias(Autodesk.Revit.DB.Document doc)
        {
            List<Category> categorias = Tools.ObtenerListaCategoriasModelo(doc);
            // Rellenar Combo
            foreach (var item in categorias)
            {
                this.cmbCategorias.Items.Add(item.Name);
            }
        }

        // Rellenar el Combo Tipos de Familias
        private void RellenarComboTiposFamilia(Autodesk.Revit.DB.Document doc, string nombreCategoria)
        {
            List<FamilySymbol> tipos = Tools.ObtenerListaTiposFamiliaModelo(doc);
            // Limpiar el combo
            this.cmbFamilias.Items.Clear();

            // Rellenar Combo
            foreach (var item in tipos)
            {
                // Verificar que sea igual a la Categoria elegida
                if (item.Category.Name == nombreCategoria)
                {
                    this.cmbFamilias.Items.Add(item.Name);
                }
            }
        }

        // Rellenar el Combo Niveles
        private void RellenarComboNiveles(Autodesk.Revit.DB.Document doc)
        {
            // Obtener la Lista de niveles
            List<Level> niveles = Tools.ObtenerListaNiveles(doc);
            // Relleno el combo
            foreach (var item in niveles)
            {
                this.cmbNiveles.Items.Add(item.Name);
            }
        }

        private void cmbCategorias_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Obtener el texto elegido
            string nombre = this.cmbCategorias.SelectedItem.ToString();

            // Activo el Combo Familias
            this.cmbFamilias.Enabled = true;

            // Rellenar el Combo Familias
            RellenarComboTiposFamilia(_doc, nombre);
        }

        // Crea una Lista de Puntos leyendo el DataGridView
        public List<XYZ> ConvertirPuntos()
        {
            // Crear lista de Puntos
            List<XYZ> listaPuntos = new List<XYZ>();
            foreach (DataGridViewRow fila in this.dgvCoordenadas.Rows)
            {
                // Verificar que no haya celdas vacias
                if (fila.Cells[0].Value != null && fila.Cells[1].Value != null)
                {
                    // Voy a leer el valor de las celdas
                    string valor0 = fila.Cells[0].Value.ToString();
                    string valor1 = fila.Cells[1].Value.ToString();
                    // Verificar si se puede convertir los valores en Numeros
                    try
                    {
                        double valorX = Convert.ToDouble(valor0);
                        double valorY = Convert.ToDouble(valor1);

                        // Conversion de unidades (Metros a Pies)
                        valorX = valorX / 0.3048;
                        valorY = valorY / 0.3048;

                        XYZ punto = new XYZ(valorX, valorY, 0);
                        // Llenar lista de puntos
                        listaPuntos.Add(punto);
                    }
                    catch (Exception)
                    {
                        // No pasa nada. No se crea ningun punto
                    }
                }
            }
            return listaPuntos;
        }

        private void btnEjemplares_Click(object sender, EventArgs e)
        {
            // Verificaciones de las selecciones del usuario
            if (this.cmbCategorias.SelectedItem == null)
            {
                TaskDialog.Show("Revit", "Debe seleccionar una Categoría");
                return;
            }
            if (this.cmbFamilias.SelectedItem == null)
            {
                TaskDialog.Show("Revit", "Debe seleccionar un Tipo de Familia");
                return;
            }
            if (this.cmbNiveles.SelectedItem == null)
            {
                TaskDialog.Show("Revit", "Debe seleccionar un Nivel");
                return;
            }
            if (this.dgvCoordenadas.Rows.Count <= 1)
            {
                TaskDialog.Show("Revit", "Debe ingresar valores de coordenadas");
                return;
            }
            // Todo Ok, procesar elementos
            // Obtener la lista de Puntos
            List<XYZ> listaPuntos = ConvertirPuntos();

            // Obtener el FamilySymbol
            string nombreFamilia = this.cmbFamilias.SelectedItem.ToString();
            FamilySymbol familia = Tools.ObtenerFamilySymbolPorNombre(_doc, nombreFamilia);

            // Obtener el Nivel seleccionado
            string nombreNivel = this.cmbNiveles.SelectedItem.ToString();
            Level nivel = Tools.ObtenerNivelPorNombre(_doc, nombreNivel);

            // Structural Type
            StructuralType tipoEstructural = StructuralType.NonStructural;

            // Crear la lista de FamilyInstanceCreationData
            List<FamilyInstanceCreationData> listaData = new List<FamilyInstanceCreationData>();

            // Recorrer los puntos XYZ
            foreach (XYZ punto in listaPuntos)
            {
                FamilyInstanceCreationData data = new FamilyInstanceCreationData(punto, familia, nivel, tipoEstructural);
                listaData.Add(data);
            }

            // Iniciar la Transaccion y crear los elementos
            Transaction t = new Transaction(_doc, "Crear elementos");
            t.Start();
            _doc.Create.NewFamilyInstances2(listaData);
            t.Commit();
            TaskDialog.Show("Revit", "Los elementos se crearon con éxito");
           
        }

        public void RellenarPuntos(List<XYZ> listaPuntos)
        {
            int fila = 0;
            foreach (var punto in listaPuntos)
            {
                // Agregar una fila nueva e insertar valores
                this.dgvCoordenadas.Rows.Add();
                this.dgvCoordenadas.Rows[fila].Cells[0].Value = punto.X;
                this.dgvCoordenadas.Rows[fila].Cells[1].Value = punto.Y;
                // Aumentar el contador
                fila++;
            }
        }

        private void btnAbrirExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivo Excel (*.xlsx)|*.xlsx";

            if (open.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string rutaArchivo = open.FileName;
                    List<XYZ> listaPuntos = Tools.ImportarPuntosExcel(rutaArchivo);
                    RellenarPuntos(listaPuntos);
                    MessageBox.Show("Puntos importados correctamente", "Importar Puntos",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: \n" + ex.Message, "Importar Puntos",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
    }
}
