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
    public partial class Historico : Form
    {
        public Historico()
        {
            InitializeComponent();
        }

        public static String rid;
        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                rid = "0";
                dataGridView1.Rows.Clear();
                if (!string.IsNullOrEmpty(buscar.Text))
                {
                    foreach (String[] activo in Form1.lista_historico)
                    {
                        String economico = activo[3].ToString().ToUpper();
                        String numserie = activo[5].ToString().ToUpper();
                        String nombre = activo[6].ToString().ToUpper();
                        if (economico.Contains(buscar.Text.ToUpper()) || numserie.Contains(buscar.Text.ToUpper()) || nombre.Contains(buscar.Text.ToUpper()))
                        {
                            dataGridView1.Rows.Add(activo[0], activo[3], activo[1], activo[4], activo[5], activo[7], activo[6]);
                        }
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
    }
}
