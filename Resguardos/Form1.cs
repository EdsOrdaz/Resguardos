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
        private String versiontext = "10.5";
        private String version = "10b3adf4529649ecb009c579e7713c8e";
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
                    SqlCommand comm = new SqlCommand("SELECT *  FROM [bd_SiRAc].[dbo].[LNJ_Equipos_Gilberto] where [No. Activo] LIKE '%" + buscar.Text.ToUpper() + "%' OR [No. Serie] LIKE '%" + buscar.Text.ToUpper() + "%' OR [Nombre de Resguardatario] LIKE '%" + buscar.Text.ToUpper() + "%'", conexion);
                    SqlDataReader nwReader = comm.ExecuteReader();
                    while (nwReader.Read())
                    {
                        String[] n = new String[18]; 
                        n[0] = nwReader["No. Activo"].ToString();
                        n[1] = nwReader["Fecha SIRAC"].ToString();
                        n[2] = nwReader["Categoría"].ToString();
                        n[3] = nwReader["Subcategoría"].ToString();
                        n[4] = nwReader["marca"].ToString();
                        n[5] = nwReader["modelo"].ToString();
                        n[6] = nwReader["No. Serie"].ToString();
                        n[7] = nwReader["importe"].ToString();
                        n[8] = nwReader["No. Factura"].ToString();
                        n[9] = nwReader["Fecha Factura"].ToString();
                        n[10] = nwReader["Ubicación"].ToString();
                        n[11] = nwReader["No. Resguardatario"].ToString();
                        n[12] = nwReader["Nombre de Resguardatario"].ToString();
                        n[13] = nwReader["Empresa"].ToString();
                        n[14] = nwReader["Centro de Costo"].ToString();
                        n[15] = nwReader["Base"].ToString();
                        n[16] = nwReader["Fecha Asignación"].ToString();
                        n[17] = nwReader["Pedido Compra"].ToString();
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
                Application.Exit();
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
                String economico = activo[0].ToString().ToUpper();
                String serie = activo[6].ToString().ToUpper();
                String nombre = activo[12].ToString().ToUpper();
                if (economico.Contains(buscar.Text.ToUpper()) || serie.Contains(buscar.Text.ToUpper()) || nombre.Contains(buscar.Text.ToUpper()))
                {
                    dataGridView1.Rows.Add(activo[0], activo[3], activo[4], activo[5], activo[6], activo[16], activo[12]);
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
                        backgroundWorker1.RunWorkerAsync();
                    }
                }
            }
        }

        private void historicoDeEquiposToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (backgroundWorker2.IsBusy != true)
            {
                label1.Visible = false;
                buscar.Visible = false;
                dataGridView1.Visible = false;
                pictureBox1.Enabled = true;
                pictureBox1.BringToFront();
                pictureBox1.Visible = true;
                menuStrip1.Visible = false;
                backgroundWorker2.RunWorkerAsync();
            }
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
                Clipboard.SetText(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
            }
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Visible = false;
            pictureBox1.Enabled = false;
            dataGridView1.Visible = true;
            label1.Visible = true;
            menuStrip1.Visible = true;
            buscar.Visible = true;

            Historico historico = new Historico();
            historico.ShowDialog();

            buscar.Focus();
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            //Actualizar lista de historico de equipos

            if (!lista_historico.Any())
            {
                try
                {
                    using (SqlConnection conexion = new SqlConnection(nsql))
                    {
                        conexion.Open();
                        SqlCommand comm = new SqlCommand("SELECT r.id_ms_resguardo, CGE.marca AS 'Marca', CGE.descripcion AS 'Descripcion', DTE.codigo as 'Num Economico', DTE.modelo as 'Modelo', DTE.no_serie as 'Num de Serie', EMP.nombre+' '+EMP.ap_paterno+' '+EMP.ap_materno AS 'Resguardado a', R.fecha_alta as 'Fecha de asignacion', EMPASIG.nombre+' '+EMPASIG.ap_paterno+' '+EMPASIG.ap_materno as 'Usuario que lo asigno', R.fecha_baja AS 'Fecha de desasignacion', EMPDESASIG.nombre+' '+EMPDESASIG.ap_paterno+' '+EMPDESASIG.ap_materno as 'Usuario que lo desasigno', R.obs AS 'Obs de resguardo', DTE.empresa_dueña as 'Empresa Dueña', UBI.descripcion AS 'Ubicacion', DTE.status AS 'Estatus del equipo',  DTE.fecha_reg AS 'Fecha de registro', EMPLEADOREGISTRO.nombre+' '+EMPLEADOREGISTRO.ap_paterno+' '+EMPLEADOREGISTRO.ap_materno AS 'Empleado que registro el equipo', DTE.pedido_compra AS 'Pedido de compra', DTE.fecha_baja AS 'Fecha de baja del equipo', EMPLEADOBAJA.nombre+' '+EMPLEADOBAJA.ap_paterno+' '+EMPLEADOBAJA.ap_materno AS 'Empleado que dio de baja el equipo', DTE.obsBaja AS 'Observaciones de Baja', DTE.obs_ti AS 'Observaciones de Sistemas'  FROM [bd_SiRAc].[dbo].[ms_resguardo] R FULL JOIN [bd_SiRAc].[dbo].[dt_equipo] DTE ON R.id_dt_equipo=DTE.id_dt_equipo FULL JOIN [bd_Empleado].[dbo].[cg_empleado] EMP ON R.id_empleado=EMP.id_empleado FULL JOIN [bd_SiRAc].[dbo].[cg_ubicacion] UBI ON DTE.id_ubicacion=UBI.id_ubicacion FULL JOIN [bd_SiRAc].[dbo].[cg_usuario] USUREGISTRO ON DTE.id_usr_reg=USUREGISTRO.id_usuario FULL JOIN [bd_Empleado].[dbo].[cg_empleado] EMPLEADOREGISTRO ON USUREGISTRO.id_empleado=EMPLEADOREGISTRO.id_empleado FULL JOIN [bd_SiRAc].[dbo].[cg_usuario] USUBAJA ON DTE.id_usr_baja=USUBAJA.id_usuario FULL JOIN [bd_Empleado].[dbo].[cg_empleado] EMPLEADOBAJA ON USUBAJA.id_empleado=EMPLEADOBAJA.id_empleado INNER JOIN [bd_SiRAc].[dbo].[cg_usuario] US ON R.id_usr_alta=US.id_usuario INNER JOIN [bd_Empleado].[dbo].[cg_empleado] EMPASIG on us.id_empleado=EMPASIG.id_empleado FULL JOIN [bd_SiRAc].[dbo].[cg_usuario] USDES ON R.id_usr_baja=USDES.id_usuario FULL JOIN [bd_Empleado].[dbo].[cg_empleado] EMPDESASIG on USDES.id_empleado=EMPDESASIG.id_empleado  INNER JOIN [bd_SiRAc].[dbo].[cg_equipo] CGEQ ON CGEQ.id_equipo=DTE.id_equipo INNER JOIN [bd_SiRAc].[dbo].[cg_categoria] CAT ON CAT.id_categoria=CGEQ.id_categoria  INNER JOIN [bd_SiRAc].[dbo].[cg_equipo] CGE ON DTE.id_equipo=CGE.id_equipo  where CAT.id_categoria=5 ORDER BY [Fecha de asignacion] DESC", conexion);
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

                            lista_historico.Add(n);
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Error en la busqueda de historico de activos.\n\nMensaje: " + ex.Message, "Información del Equipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
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
                            Application.Exit();
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se puede verificar la version del programa.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error en validar la version del programa\n\nMensaje: " + ex.Message, "Información del Equipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }
        }

        private void RevisarVersion_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                Application.Exit();
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
    }
}
