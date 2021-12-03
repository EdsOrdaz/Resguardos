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
                String economico = activo[0].ToString().ToUpper();
                if (economico == Form1.numeconomico)
                {
                    dataGridView1.Rows.Add("No. Activo:", activo[0]);
                    dataGridView1.Rows.Add("Fecha SIRAC:", activo[1]);
                    dataGridView1.Rows.Add("Categoría:", activo[2]);
                    dataGridView1.Rows.Add("Subcategoría:", activo[3]);
                    dataGridView1.Rows.Add("marca:", activo[4]);
                    dataGridView1.Rows.Add("modelo:", activo[5]);
                    dataGridView1.Rows.Add("No. Serie:", activo[6]);
                    dataGridView1.Rows.Add("importe:", activo[7]);
                    dataGridView1.Rows.Add("No. Factura:", activo[8]);
                    dataGridView1.Rows.Add("Fecha Factura:", activo[9]);
                    dataGridView1.Rows.Add("Ubicación:", activo[10]);
                    dataGridView1.Rows.Add("No. Resguardatario:", activo[11]);
                    dataGridView1.Rows.Add("Nombre de Resguardatario:", activo[12]);
                    dataGridView1.Rows.Add("Empresa:", activo[13]);
                    dataGridView1.Rows.Add("Centro de Costo:", activo[14]);
                    dataGridView1.Rows.Add("Base:", activo[15]);
                    dataGridView1.Rows.Add("Fecha Asignación:", activo[16]);
                    dataGridView1.Rows.Add("Pedido Compra:", activo[17]);


                    System.Windows.Forms.DataGridViewCellStyle boldStyle = new System.Windows.Forms.DataGridViewCellStyle();
                    boldStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);

                    for (int i = 0; i <= 17; i++)
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
