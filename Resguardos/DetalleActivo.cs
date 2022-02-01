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
    public partial class DetalleActivo : Form
    {
        public DetalleActivo()
        {
            InitializeComponent();
        }

        private void DetalleActivo_Load(object sender, EventArgs e)
        {
            foreach (String[] activo in Form1.lista)
            {
                String economico = activo[4].ToString().ToUpper();
                if (economico == Form1.numeconomico)
                {
                    dataGridView1.Rows.Add("Categoría", activo[0]);
                    dataGridView1.Rows.Add("Subcategoría", activo[1]);
                    dataGridView1.Rows.Add("No. Activo", activo[4]);
                    dataGridView1.Rows.Add("Marca", activo[24]);
                    dataGridView1.Rows.Add("Modelo", activo[25]);
                    dataGridView1.Rows.Add("importe", activo[5]);
                    dataGridView1.Rows.Add("No. Factura", activo[6]);
                    dataGridView1.Rows.Add("Ubicación", activo[7]);
                    dataGridView1.Rows.Add("Num. Empleado", activo[8]);
                    dataGridView1.Rows.Add("Nombre de Resguardatario", activo[9]);
                    dataGridView1.Rows.Add("CC", activo[10]);
                    dataGridView1.Rows.Add("Observaciones", activo[11]);
                    dataGridView1.Rows.Add("Fecha Asignación", activo[12]);
                    dataGridView1.Rows.Add("capitalizable", activo[13]);
                    dataGridView1.Rows.Add("No. Serie", activo[14]);
                    dataGridView1.Rows.Add("Empresa Dueña", activo[15]);
                    dataGridView1.Rows.Add("Usuario Reg. Equipo", activo[16]);
                    dataGridView1.Rows.Add("Fecha Reg. Equipo", activo[17]);
                    dataGridView1.Rows.Add("Usuario Mod. Equipo", activo[18]);
                    dataGridView1.Rows.Add("Fecha Mod. Equipo", activo[19]);
                    dataGridView1.Rows.Add("Usuario Alta Resg.", activo[20]);
                    dataGridView1.Rows.Add("Fecha Alta Resg.", activo[21]);
                    dataGridView1.Rows.Add("Usuario Baja Resg.", activo[22]);
                    dataGridView1.Rows.Add("Fecha Baja Resg.", activo[23]);
                    dataGridView1.Rows.Add("No. Parte", activo[2]);
                    dataGridView1.Rows.Add("medida", activo[3]);



                    System.Windows.Forms.DataGridViewCellStyle boldStyle = new System.Windows.Forms.DataGridViewCellStyle();
                    boldStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);

                    for (int i = 0; i <= 25; i++)
                    {
                        dataGridView1.Rows[i].Cells[0].Style = boldStyle;
                        dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Silver;
                    }
                    this.dataGridView1.ClearSelection();
                }
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Clipboard.SetText(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
        }
    }
}
