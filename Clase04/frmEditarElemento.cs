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
    public partial class frmEditarElemento : System.Windows.Forms.Form
    {
        public Document _doc;
        public Element _elem;

        public frmEditarElemento(Document doc, Element elem)
        {
            InitializeComponent();
            _elem = elem;
            _doc = doc;
            // Rellenar informacion
            this.txtNombre.Text = elem.Name;
            this.txtComentario.Text = elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString();
            this.txtMarca.Text = elem.get_Parameter(BuiltInParameter.DOOR_NUMBER).AsString();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // Obtener valores de Parametros
            string comentario = this.txtComentario.Text;
            string marca = this.txtMarca.Text;

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

            this.Close();
        }
    }
}
