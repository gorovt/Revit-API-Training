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

namespace AddinHolaMundo
{
    public partial class frmColecciones : System.Windows.Forms.Form
    {
        public Document _doc;

        public frmColecciones(Document doc, List<Element> lista)
        {
            InitializeComponent();
            _doc = doc;
            FillTree(lista);
        }

        private void FillTree(List<Element> lista)
        {
            // Crear Nodo principal
            TreeNode node0 = new TreeNode();
            node0.Name = "00";
            node0.Text = "Lista de Muros:";

            // Recorrer la Lista y crear un Nodo para cada elemento
            foreach (Element elem in lista)
            {
                TreeNode node1 = new TreeNode();
                node1.Name = elem.Id.ToString();
                node1.Text = elem.Name + "<" + elem.Id.ToString() + ">";
                node0.Nodes.Add(node1);
            }

            // Colocar el Nodo principal en el Arbol
            this.trvMuros.Nodes.Add(node0);
        }

        private void FillParametros(Element elem)
        {
            // Crear un StringBuilder
            StringBuilder sb = new StringBuilder();

            // Agregar Titulo Ejemplar
            sb.AppendLine("*** Parámetros de Ejemplar ***");

            // Obtener todos los parametros de Ejemplar del elemento
            foreach (Parameter item in elem.Parameters)
            {
                sb.AppendLine(item.Definition.Name + ": " + item.AsValueString());
            }

            // Agregar Titulo Tipo
            sb.AppendLine("*** Parámetros de Tipo ***");

            // Obtener el elemento de Tipo de Familia
            ElementId id = elem.GetTypeId();
            Element tipo = _doc.GetElement(id);

            // Obtener todos los parametros de Tipo del elemento
            foreach (Parameter item in tipo.Parameters)
            {
                sb.AppendLine(item.Definition.Name + ": " + item.AsValueString());
            }

            // Completar Textos
            this.txtParametros.Text = sb.ToString();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void trvMuros_AfterSelect(object sender, TreeViewEventArgs e)
        {
            this.txtParametros.Text = string.Empty;
            string nombre = e.Node.Name;
            if (nombre == "00")
            {
                this.txtParametros.Text = "Seleccione un elemento de la Lista";
                this.dgvParametros.DataBindings.Clear();
                this.dgvParametros.DataSource = new List<ParamGrid>();
            }
            else
            {
                // Completar el RichTextBox
                int id = Convert.ToInt32(nombre);
                ElementId elemId = new ElementId(id);
                Element selected = _doc.GetElement(elemId);
                FillParametros(selected);

                // Completar el DataGridView
                FillDataGrid(selected);
            }
        }

        private void FillDataGrid(Element elem)
        {
            // Crear una lista de ParamGrid
            List<ParamGrid> lista = new List<ParamGrid>();
                 
            // Obtener todos los parametros de Ejemplar del elemento
            foreach (Parameter item in elem.Parameters)
            {
                // Crear un objeto ParamGrid
                int id = item.Id.IntegerValue;
                string nombre = item.Definition.Name;
                string valor = item.AsString();
                if (valor == null)
                {
                    valor = item.AsValueString();
                }
                ParamGrid param = new ParamGrid(id, nombre, valor);
                lista.Add(param);
            }

            this.dgvParametros.DataBindings.Clear();
            this.dgvParametros.DataSource = lista;
        }
    }
}
