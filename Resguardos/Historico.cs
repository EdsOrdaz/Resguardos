using Microsoft.Office.Interop.Excel;
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
                    foreach (DataGridViewColumn column in dataGridView1.Columns)
                        if (column.Visible)
                            hojaTrabajo.Cells[1, ++iCol] = column.HeaderText;
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
                                hojaTrabajo.Cells[i + 2, j + 1] = grd.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                    }
                    librosTrabajo.SaveAs(fichero.FileName, XlFileFormat.xlWorkbookNormal,
                                          System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false,
                                          XlSaveAsAccessMode.xlShared, false, false,
                                          System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
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
    }
}
