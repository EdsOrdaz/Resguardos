using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Resguardos
{
    public partial class Detalle_Historico : Form
    {
        public Detalle_Historico()
        {
            InitializeComponent();
        }

        private void Detalle_Historico_Load(object sender, EventArgs e)
        {
            foreach (String[] activo in Historico.lista)
            {
                String mrid = activo[0].ToString().ToUpper();
                if (mrid == Historico.rid)
                {
                    dataGridView1.Rows.Add("id_ms_resguardo", activo[0]);
                    dataGridView1.Rows.Add("Marca", activo[1]);
                    dataGridView1.Rows.Add("Descripcion", activo[2]);
                    dataGridView1.Rows.Add("Num Economico", activo[3]);
                    dataGridView1.Rows.Add("Modelo", activo[4]);
                    dataGridView1.Rows.Add("Num de Serie", activo[5]);
                    dataGridView1.Rows.Add("Resguardado a", activo[6]);
                    dataGridView1.Rows.Add("Fecha de asignacion", activo[7]);
                    dataGridView1.Rows.Add("Usuario que lo asigno", activo[8]);
                    dataGridView1.Rows.Add("Fecha de desasignacion", activo[9]);
                    dataGridView1.Rows.Add("Usuario que lo desasigno", activo[10]);
                    dataGridView1.Rows.Add("Obs de resguardo", activo[11]);
                    dataGridView1.Rows.Add("Empresa Dueña", activo[12]);
                    dataGridView1.Rows.Add("Ubicacion", activo[13]);
                    dataGridView1.Rows.Add("Estatus del equipo", activo[14]);
                    dataGridView1.Rows.Add("Fecha de registro", activo[15]);
                    dataGridView1.Rows.Add("Empleado que registro el equipo", activo[16]);
                    dataGridView1.Rows.Add("Pedido de compra", activo[17]);
                    dataGridView1.Rows.Add("Fecha de baja del equipo", activo[18]);
                    dataGridView1.Rows.Add("Empleado que dio de baja el equipo", activo[19]);
                    dataGridView1.Rows.Add("Observaciones de Baja", activo[20]);
                    dataGridView1.Rows.Add("Observaciones de Sistemas", activo[21]);


                    System.Windows.Forms.DataGridViewCellStyle boldStyle = new System.Windows.Forms.DataGridViewCellStyle();
                    boldStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);

                    for (int i = 0; i <= 21; i++)
                    {
                        dataGridView1.Rows[i].Cells[0].Style = boldStyle;
                        dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Silver;
                    }
                    this.dataGridView1.ClearSelection();
                }
            }
        }
    }
}
