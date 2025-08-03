using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProyectodeTeoríadelaComputación
{
    public class FormatoCeldas
    {
        private DataGridView DGVCeldas;

        public FormatoCeldas(DataGridView DGVCeldas)
        {
            this.DGVCeldas = DGVCeldas;
        }

        public void AplicarNegrita()
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                Font currentFont = cell.Style.Font ?? DGVCeldas.DefaultCellStyle.Font ?? DGVCeldas.Font;
                FontStyle newStyle = currentFont.Style ^ FontStyle.Bold;
                cell.Style.Font = new Font(currentFont, newStyle);
            }
        }

        public void AplicarCursiva()
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                Font currentFont = cell.Style.Font ?? DGVCeldas.DefaultCellStyle.Font ?? DGVCeldas.Font;
                FontStyle newStyle = currentFont.Style ^ FontStyle.Italic;
                cell.Style.Font = new Font(currentFont, newStyle);
            }
        }

        public void AplicarSubrayado()
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                Font currentFont = cell.Style.Font ?? DGVCeldas.DefaultCellStyle.Font ?? DGVCeldas.Font;
                FontStyle newStyle = currentFont.Style ^ FontStyle.Underline;
                cell.Style.Font = new Font(currentFont, newStyle);
            }
        }

        public void CambiarFuente(string fuente)
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                Font currentFont = cell.Style.Font ?? DGVCeldas.DefaultCellStyle.Font ?? DGVCeldas.Font;
                cell.Style.Font = new Font(fuente, currentFont.Size, currentFont.Style);
            }
        }

        public void CambiarTamaño(float tamaño)
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                Font currentFont = cell.Style.Font ?? DGVCeldas.DefaultCellStyle.Font ?? DGVCeldas.Font;
                cell.Style.Font = new Font(currentFont.FontFamily, tamaño, currentFont.Style);
            }
        }

        public void AplicarRellenoColor(Color color)
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                cell.Style.BackColor = color;
            }
        }

        public void AplicarColorTexto(Color color)
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                cell.Style.ForeColor = color;
            }
        }

        public void AutoAjustarColumnas()
        {
            DGVCeldas.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            int minWidth = 80;
            foreach (DataGridViewColumn col in DGVCeldas.Columns)
            {
                if (col.Width < minWidth)
                    col.Width = minWidth;
            }
        }

        public void AutoAjustarFilas()
        {
            DGVCeldas.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);

            int minHeight = 22;
            foreach (DataGridViewRow row in DGVCeldas.Rows)
            {
                if (row.Height < minHeight)
                    row.Height = minHeight;
            }
        }

        public void AutoAjustarTodo()
        {
            AutoAjustarColumnas();
            AutoAjustarFilas();
        }

        public void LimpiarCeldas(TextBox barraFormulas)
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                cell.Value = null;
                cell.Style.BackColor = DGVCeldas.DefaultCellStyle.BackColor;
                cell.Style.ForeColor = DGVCeldas.DefaultCellStyle.ForeColor;
                cell.Style.Font = DGVCeldas.DefaultCellStyle.Font;

                var key = (cell.ColumnIndex, cell.RowIndex);
                ReferenciasCeldas.RemoverFormula(key);
            }

            barraFormulas.Text = "";
        }
    }
}
