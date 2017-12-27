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
using Autodesk.Revit.DB.Architecture;

namespace AddinHolaMundo
{
    public partial class frmFm : System.Windows.Forms.Form
    {
        public Document _doc;
        public frmFm(Document doc, List<Room> lstRooms, List<FamilyInstance> lstFamilies)
        {
            InitializeComponent();
            _doc = doc;
            fillTreeView(lstRooms, lstFamilies);
        }

        private void fillTreeView(List<Room> lstRooms, List<FamilyInstance> lstFamilies)
        {
            TreeNode node0 = new TreeNode();
            node0.Name = "Habitaciones";
            node0.Text = "Habitaciones";
            node0.ExpandAll();

            // Habitaciones
            foreach (var item in lstRooms)
            {
                TreeNode node1 = new TreeNode();
                node1.Name = item.Id.ToString();
                node1.Text = item.Name;

                foreach (var familia in lstFamilies)
                {
                    if (familia.Room != null && familia.Room.Id == item.Id)
                    {
                        TreeNode node2 = new TreeNode();
                        node2.Name = familia.Id.ToString();
                        node2.Text = familia.Name + "<" + familia.Id + ">";
                        node1.Nodes.Add(node2);
                    }
                }
                node0.Nodes.Add(node1);
            }

            this.trvHabitaciones.Nodes.Add(node0);
        }

        private void fillProperties(Element elem)
        {

        }

        private void trvHabitaciones_AfterSelect(object sender, TreeViewEventArgs e)
        {
            this.prpElemento.SelectedObject = null;

            if (e.Node.Level == 1)
            {
                // Se trata de una Habitación
                string id = e.Node.Name;
                ElementId elemId = new ElementId(Convert.ToInt32(id));
                Element elem = _doc.GetElement(elemId);
                Room hab = elem as Room;
                this.prpElemento.SelectedObject = hab;
            }

            if (e.Node.Level == 2)
            {
                // Se trata de una familia
                string id = e.Node.Name;
                ElementId elemId = new ElementId(Convert.ToInt32(id));
                Element elem = _doc.GetElement(elemId);
                FamilyInstance fam = elem as FamilyInstance;
                this.prpElemento.SelectedObject = fam;
            }
        }

        private void btnExportCsv_Click(object sender, EventArgs e)
        {
            string rutaArchivo = string.Empty;
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Archivo CSV (*.csv)|*.csv";
            string title = "Habitaciones";
            sfd.FileName = title + ".csv";
            DialogResult result = sfd.ShowDialog();
            if (result == DialogResult.OK)
            {
                rutaArchivo = sfd.FileName;
                try
                {
                    Tools.ExportarTreeViewCsv(this.trvHabitaciones, rutaArchivo);
                    MessageBox.Show("Se exportaron las habitaciones correctamente", "Habitaciones",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message, "Habitaciones",
                       MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
    }
}
