using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Resguardos
{
    public partial class Form1 : Form
    {
        /*
         V10.7.1
        - Se amplia la busqueda a todas las categorias
        - Se agrega opcion de copiar informacion del equipo con la letra C
        */

        private String versiontext = "10.7.1";
        private String version = "e32955e68c67c2be6bbc7567f3e9601f";
        public static String conexionsqllast = "server=148.223.153.37,5314; database=InfEq;User ID=eordazs;Password=Corpame*2013; integrated security = false ; MultipleActiveResultSets=True";

        public static String servivor = "148.223.153.43\\MSSQLSERVER1";
        public static String basededatos = "bd_SiRAc";
        public static String usuariobd = "sa";
        public static String passbd = "At3n4";
        public static string nsql = "server=" + servivor + "; database=" + basededatos + " ;User ID=" + usuariobd + ";Password=" + passbd + "; integrated security = false ; MultipleActiveResultSets=True";


        public static List<String[]> lista = new List<String[]>();
        public static List<String[]> lista_historico = new List<String[]>();

        public static String numeconomico;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            lista.Clear();
            lista_historico.Clear();
            dataGridView1.Rows.Clear();
            menuStrip1.Visible = false;
            label1.Visible = false;
            label2.Visible = false;

            this.FormBorderStyle = FormBorderStyle.None;
            this.TransparencyKey = Color.Gray;
            this.BackColor = Color.Gray;

            if (RevisarVersion.IsBusy != true)
            {
                buscar.Visible = false;
                dataGridView1.Visible = false;
                pictureBox1.Enabled = true;
                pictureBox1.BringToFront();
                pictureBox1.Visible = true;
                RevisarVersion.RunWorkerAsync();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            //Actualizar lista de resguardos actuales
            try
            {
                using (SqlConnection conexion = new SqlConnection(nsql))
                {
                    conexion.Open();
                    SqlCommand comm = new SqlCommand("SELECT e.*, cg.marca, dt.modelo FROM [bd_SiRAc].[dbo].[LNJ_Equipos] e " +
                        "INNER JOIN [bd_SiRAc].[dbo].[dt_equipo] dt ON dt.codigo=e.[No. Activo] " +
                        "INNER JOIN [bd_SiRAc].[dbo].[LNJ_cg_equipo] cg ON dt.id_equipo=cg.id_equipo" +
                        " where e.[No. Activo] LIKE '%" + buscar.Text.ToUpper() + "%' OR e.[No. Serie] LIKE '%" + buscar.Text.ToUpper() + "%' OR e.[Nombre de Resguardatario] LIKE '%" + buscar.Text.ToUpper() + "%'", conexion);
                    SqlDataReader nwReader = comm.ExecuteReader();
                    while (nwReader.Read())
                    {
                        String[] n = new String[26]; 
                        n[0] = nwReader["Categoría"].ToString();
                        n[1] = nwReader["Subcategoría"].ToString();
                        n[2] = nwReader["No. Parte"].ToString();
                        n[3] = nwReader["medida"].ToString();
                        n[4] = nwReader["No. Activo"].ToString();
                        n[5] = nwReader["importe"].ToString();
                        n[6] = nwReader["No. Factura"].ToString();
                        n[7] = nwReader["Ubicación"].ToString();
                        n[8] = nwReader["Num. Empleado"].ToString();
                        n[9] = nwReader["Nombre de Resguardatario"].ToString();
                        n[10] = nwReader["CC"].ToString();
                        n[11] = nwReader["Observaciones"].ToString();
                        n[12] = nwReader["Fecha Asignación"].ToString();
                        n[13] = nwReader["capitalizable"].ToString();
                        n[14] = nwReader["No. Serie"].ToString();
                        n[15] = nwReader["Empresa Dueña"].ToString();
                        n[16] = nwReader["Usuario Reg. Equipo"].ToString();
                        n[17] = nwReader["Fecha Reg. Equipo"].ToString();
                        n[18] = nwReader["Usuario Mod. Equipo"].ToString();
                        n[19] = nwReader["Fecha Mod. Equipo"].ToString();
                        n[20] = nwReader["Usuario Alta Resg."].ToString();
                        n[21] = nwReader["Fecha Alta Resg."].ToString();
                        n[22] = nwReader["Usuario Baja Resg."].ToString();
                        n[23] = nwReader["Fecha Baja Resg."].ToString();
                        n[24] = nwReader["marca"].ToString();
                        n[25] = nwReader["modelo"].ToString();

                        lista.Add(n);
                    }
                }
            }
            catch (System.Exception ex)
            {
                e.Cancel = true;
                MessageBox.Show("Error en la busqueda\n\nMensaje: " + ex.Message, "Información del Equipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                System.Windows.Forms.Application.Exit();
            }
            pictureBox1.Visible = false;
            pictureBox1.Enabled = false;
            dataGridView1.Visible = true;

            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.BackColor = Control.DefaultBackColor;
            label1.Visible = true;
            this.Icon = Resguardos.Properties.Resources.Resguardos;

            menuStrip1.Visible = true;
            buscar.Visible = true;
            buscar.Focus();

            foreach (String[] activo in lista)
            {
                String economico = activo[4].ToString().ToUpper();
                String serie = activo[14].ToString().ToUpper();
                String nombre = activo[9].ToString().ToUpper();

                if (economico.Contains(buscar.Text.ToUpper()) || serie.Contains(buscar.Text.ToUpper()) || nombre.Contains(buscar.Text.ToUpper()))
                {
                    dataGridView1.Rows.Add(activo[4], activo[1], activo[24], activo[25], activo[14], activo[12], activo[9]);
                }
            }
        }

        private void eco_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dataGridView1.Rows.Clear();
                if (!string.IsNullOrEmpty(buscar.Text))
                {

                    lista.Clear();
                    lista_historico.Clear();
                    dataGridView1.Rows.Clear();
                    menuStrip1.Visible = false;
                    label1.Visible = false;

                    if (backgroundWorker1.IsBusy != true)
                    {
                        buscar.Visible = false;
                        dataGridView1.Visible = false;
                        pictureBox1.Enabled = true;
                        pictureBox1.BringToFront();
                        pictureBox1.Visible = true;
                        label2.Visible = false;
                        backgroundWorker1.RunWorkerAsync();
                    }
                }
            }
        }

        private void historicoDeEquiposToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Historico historico = new Historico();
            historico.ShowDialog();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                numeconomico = dataGridView1.Rows[e.RowIndex].Cells["economico"].Value.ToString();
                DetalleActivo detalle = new DetalleActivo();
                detalle.ShowDialog();
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                label2.Visible = true;
                label2.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                Clipboard.SetText(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
            }
        }

        private void buscar_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            buscar.Text= Clipboard.GetText(TextDataFormat.Text);
        }

        private void RevisarVersion_DoWork(object sender, DoWorkEventArgs e)
        {
            //Revisar version del programa
            try
            {
                using (SqlConnection conexion2 = new SqlConnection(conexionsqllast))
                {
                    conexion2.Open();
                    String sql2 = "SELECT (SELECT valor FROM Configuracion WHERE nombre='Resguardos_Version') as version,valor FROM Configuracion WHERE nombre='Resguardos_hash'";
                    SqlCommand comm2 = new SqlCommand(sql2, conexion2);
                    SqlDataReader nwReader2 = comm2.ExecuteReader();
                    if (nwReader2.Read())
                    {
                        if (nwReader2["valor"].ToString() != version)
                        {
                            MessageBox.Show("No se puede inciar porque hay una nueva version.\n\nNueva Version: " + nwReader2["version"].ToString() + "\nVersion actual: " + versiontext + "\n\nEl programa se cerrara.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            System.Windows.Forms.Application.Exit();
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se puede verificar la version del programa.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        System.Windows.Forms.Application.Exit();
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error en validar la version del programa\n\nMensaje: " + ex.Message, "Información del Equipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
                return;
            }
        }

        private void RevisarVersion_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                System.Windows.Forms.Application.Exit();
            }
            pictureBox1.Visible = false;
            pictureBox1.Enabled = false;
            dataGridView1.Visible = true;

            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.BackColor = Control.DefaultBackColor;
            label1.Visible = true;
            this.Icon = Resguardos.Properties.Resources.Resguardos;

            menuStrip1.Visible = true;
            buscar.Visible = true;
            buscar.Focus();
        }

        void ExportarDataGridViewExcel(DataGridView grd)
        {
            using (SaveFileDialog fichero = new SaveFileDialog { Filter = @"Excel (*.xls)|*.xls", FileName=buscar.Text })
            {
                if (fichero.ShowDialog() == DialogResult.OK)
                {
                    Microsoft.Office.Interop.Excel.Application aplicacion = new Microsoft.Office.Interop.Excel.Application();
                    Workbook librosTrabajo = aplicacion.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                    Worksheet hojaTrabajo = (Worksheet)librosTrabajo.Worksheets.get_Item(1);
                    int iCol = 0;
                    int iCol2 = 0;
                    foreach (DataGridViewColumn column in dataGridView1.Columns)
                    {
                        if (column.Visible)
                        {
                            ++iCol;
                            iCol2 = iCol;
                            hojaTrabajo.Cells[1, iCol2].Borders.LineStyle = XlLineStyle.xlContinuous;
                            hojaTrabajo.Cells[1, iCol2].EntireRow.Font.Bold = true;
                            hojaTrabajo.Cells[1, iCol2].Interior.Color = XlRgbColor.rgbGray;
                            hojaTrabajo.Cells[1, iCol2].Font.Color = Color.White;
                            hojaTrabajo.Cells[1, iCol2] = column.HeaderText;
                        }
                    }
                    for (int i = 0; i < grd.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < grd.Columns.Count; j++)
                        {
                            if (grd.Rows[i].Cells[j].Value is null)
                            {
                                hojaTrabajo.Cells[i + 2, j + 1] = "";
                            }
                            else
                            {
                                hojaTrabajo.Cells[i + 2, j + 1].Borders.LineStyle = XlLineStyle.xlContinuous;
                                hojaTrabajo.Cells[i + 2, j + 1] = grd.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                    }
                    aplicacion.ActiveWindow.DisplayGridlines = false;
                    librosTrabajo.SaveAs(fichero.FileName, XlFileFormat.xlWorkbookNormal);
                    aplicacion.Quit();
                    MessageBox.Show("Resguardo Guardado.", "Resguardos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void exportarAExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(buscar.Text))
            {
                ExportarDataGridViewExcel(dataGridView1);
            }
        }

        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.KeyCode==Keys.C)
            {
                String tipo = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                String economico = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                String marca = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                String modelo = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                String serie = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                label2.Text = "Información del equipo copiada al portapapeles.";
                Clipboard.SetText("Equipo: "+tipo+"\nEconomico: "+economico+"\nMarca: "+marca+"\nModelo: "+modelo+"\nNúm. Serie: "+serie);
            }
        }
    }
}
