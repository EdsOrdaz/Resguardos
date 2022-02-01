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
    public partial class Historico : Form
    {
        public Historico()
        {
            InitializeComponent();
        }

        String categoria = "";
        public static List<String[]> lista = new List<String[]>();
        public static String rid;

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dataGridView1.Rows.Clear();
                if (!string.IsNullOrEmpty(buscar.Text))
                {

                    lista.Clear();
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
                        backgroundWorker1.RunWorkerAsync();
                    }
                }
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                Clipboard.SetText(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
            }
        }

        private void Historico_Load(object sender, EventArgs e)
        {
            pictureBox1.Visible = false;
            pictureBox1.Enabled = false;
            buscar.Focus();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                rid = dataGridView1.Rows[e.RowIndex].Cells["id"].Value.ToString();
                Detalle_Historico detalle = new Detalle_Historico();
                detalle.ShowDialog();
                buscar.Focus();
            }
        }

        private void buscar_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            buscar.Text = Clipboard.GetText(TextDataFormat.Text);
        }

        void ExportarDataGridViewExcel(DataGridView grd)
        {
            using (SaveFileDialog fichero = new SaveFileDialog { Filter = @"Excel (*.xls)|*.xls", FileName = buscar.Text })
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

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                using (SqlConnection conexion = new SqlConnection(Form1.nsql))
                {
                    conexion.Open();
                    //String categoria = "CAT.id_categoria=5 AND";

                    SqlCommand comm = new SqlCommand("SELECT r.id_ms_resguardo, CGE.marca AS 'Marca', CGE.descripcion AS 'Descripcion', DTE.codigo as 'Num Economico', DTE.modelo as 'Modelo', DTE.no_serie as 'Num de Serie', EMP.nombre+' '+EMP.ap_paterno+' '+EMP.ap_materno AS 'Resguardado a', R.fecha_alta as 'Fecha de asignacion', EMPASIG.nombre+' '+EMPASIG.ap_paterno+' '+EMPASIG.ap_materno as 'Usuario que lo asigno', R.fecha_baja AS 'Fecha de desasignacion', EMPDESASIG.nombre+' '+EMPDESASIG.ap_paterno+' '+EMPDESASIG.ap_materno as 'Usuario que lo desasigno', R.obs AS 'Obs de resguardo', DTE.empresa_dueña as 'Empresa Dueña', UBI.descripcion AS 'Ubicacion', DTE.status AS 'Estatus del equipo',  DTE.fecha_reg AS 'Fecha de registro', EMPLEADOREGISTRO.nombre+' '+EMPLEADOREGISTRO.ap_paterno+' '+EMPLEADOREGISTRO.ap_materno AS 'Empleado que registro el equipo', DTE.pedido_compra AS 'Pedido de compra', DTE.fecha_baja AS 'Fecha de baja del equipo', EMPLEADOBAJA.nombre+' '+EMPLEADOBAJA.ap_paterno+' '+EMPLEADOBAJA.ap_materno AS 'Empleado que dio de baja el equipo', DTE.obsBaja AS 'Observaciones de Baja', DTE.obs_ti AS 'Observaciones de Sistemas'  FROM [bd_SiRAc].[dbo].[ms_resguardo] R FULL JOIN [bd_SiRAc].[dbo].[dt_equipo] DTE ON R.id_dt_equipo=DTE.id_dt_equipo FULL JOIN [bd_Empleado].[dbo].[cg_empleado] EMP ON R.id_empleado=EMP.id_empleado FULL JOIN [bd_SiRAc].[dbo].[cg_ubicacion] UBI ON DTE.id_ubicacion=UBI.id_ubicacion FULL JOIN [bd_SiRAc].[dbo].[cg_usuario] USUREGISTRO ON DTE.id_usr_reg=USUREGISTRO.id_usuario FULL JOIN [bd_Empleado].[dbo].[cg_empleado] EMPLEADOREGISTRO ON USUREGISTRO.id_empleado=EMPLEADOREGISTRO.id_empleado FULL JOIN [bd_SiRAc].[dbo].[cg_usuario] USUBAJA ON DTE.id_usr_baja=USUBAJA.id_usuario FULL JOIN [bd_Empleado].[dbo].[cg_empleado] EMPLEADOBAJA ON USUBAJA.id_empleado=EMPLEADOBAJA.id_empleado INNER JOIN [bd_SiRAc].[dbo].[cg_usuario] US ON R.id_usr_alta=US.id_usuario INNER JOIN [bd_Empleado].[dbo].[cg_empleado] EMPASIG on us.id_empleado=EMPASIG.id_empleado FULL JOIN [bd_SiRAc].[dbo].[cg_usuario] USDES ON R.id_usr_baja=USDES.id_usuario FULL JOIN [bd_Empleado].[dbo].[cg_empleado] EMPDESASIG on USDES.id_empleado=EMPDESASIG.id_empleado  INNER JOIN [bd_SiRAc].[dbo].[cg_equipo] CGEQ ON CGEQ.id_equipo=DTE.id_equipo INNER JOIN [bd_SiRAc].[dbo].[cg_categoria] CAT ON CAT.id_categoria=CGEQ.id_categoria  INNER JOIN [bd_SiRAc].[dbo].[cg_equipo] CGE ON DTE.id_equipo=CGE.id_equipo " +
                        "where "+categoria+" (EMP.nombre+' '+EMP.ap_paterno+' '+EMP.ap_materno LIKE '%"+buscar.Text+ "%' OR dte.codigo LIKE '%" + buscar.Text + "%' OR dte.no_serie LIKE '%" + buscar.Text + "%')" +
                        "ORDER BY [Fecha de asignacion] DESC", conexion);
                    SqlDataReader nwReader = comm.ExecuteReader();
                    while (nwReader.Read())
                    {
                        String[] n = new String[22];
                        n[0] = nwReader["id_ms_resguardo"].ToString();
                        n[1] = nwReader["Marca"].ToString();
                        n[2] = nwReader["Descripcion"].ToString();
                        n[3] = nwReader["Num Economico"].ToString();
                        n[4] = nwReader["Modelo"].ToString();
                        n[5] = nwReader["Num de Serie"].ToString();
                        n[6] = nwReader["Resguardado a"].ToString();
                        n[7] = nwReader["Fecha de asignacion"].ToString();
                        n[8] = nwReader["Usuario que lo asigno"].ToString();
                        n[9] = nwReader["Fecha de desasignacion"].ToString();
                        n[10] = nwReader["Usuario que lo desasigno"].ToString();
                        n[11] = nwReader["Obs de resguardo"].ToString();
                        n[12] = nwReader["Empresa Dueña"].ToString();
                        n[13] = nwReader["Ubicacion"].ToString();
                        n[14] = nwReader["Estatus del equipo"].ToString();
                        n[15] = nwReader["Fecha de registro"].ToString();
                        n[16] = nwReader["Empleado que registro el equipo"].ToString();
                        n[17] = nwReader["Pedido de compra"].ToString();
                        n[18] = nwReader["Fecha de baja del equipo"].ToString();
                        n[19] = nwReader["Empleado que dio de baja el equipo"].ToString();
                        n[20] = nwReader["Observaciones de Baja"].ToString();
                        n[21] = nwReader["Observaciones de Sistemas"].ToString();

                        lista.Add(n);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error en la busqueda de historico de activos.\n\nMensaje: " + ex.Message, "Información del Equipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Visible = false;
            pictureBox1.Enabled = false;
            menuStrip1.Visible = true;
            label1.Visible = true;
            buscar.Visible = true;
            dataGridView1.Visible = true;

            foreach (String[] activo in lista)
            {
                String economico = activo[3].ToString().ToUpper();
                String serie = activo[5].ToString().ToUpper();
                String nombre = activo[6].ToString().ToUpper();

                if (economico.Contains(buscar.Text.ToUpper()) || serie.Contains(buscar.Text.ToUpper()) || nombre.Contains(buscar.Text.ToUpper()))
                {
                    dataGridView1.Rows.Add(activo[0], activo[3], activo[1], activo[4], activo[5], activo[7], activo[6]);
                }
            }
        }
    }
}
