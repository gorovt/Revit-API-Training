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

namespace AddinHolaMundoTest
{
    public partial class frmSeleccionarElemento : System.Windows.Forms.Form
    {
        public frmSeleccionarElemento(Element e)
        {
            InitializeComponent();

            // Crear un elemento StringBuilder
            StringBuilder sb = new StringBuilder();

            // Recorrer todos los Parámetros del elemento
            foreach (Parameter param in e.Parameters)
            {
                sb.AppendLine(param.Definition.Name + ": " + param.AsValueString());
            }

            // Publicar los resultados en pantalla
            this.txtNombre.Text = e.Name;
            this.txtDetalles.Text = sb.ToString();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
