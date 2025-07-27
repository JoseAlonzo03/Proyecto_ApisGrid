using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProyectodeTeoríadelaComputación
{
    public class GestorTabla
    {
        private DataGridView DGVCeldas;

        public GestorTabla(DataGridView DGVCeldas)
        {
            this.DGVCeldas = DGVCeldas;
        }

        public void CopiarPegar(KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                e.Handled = true;

                string clipboardText = Clipboard.GetText();

                if (string.IsNullOrEmpty(clipboardText))
                    return;

                int startRow = DGVCeldas.CurrentCell.RowIndex;
                int startCol = DGVCeldas.CurrentCell.ColumnIndex;

                string[] lines = clipboardText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                for (int i = 0; i < lines.Length; i++)
                {
                    string[] cells = lines[i].Split('\t');

                    for (int j = 0; j < cells.Length; j++)
                    {
                        int rowIndex = startRow + i;
                        int colIndex = startCol + j;

                        while (DGVCeldas.Rows.Count <= rowIndex)
                            DGVCeldas.Rows.Add();

                        while (DGVCeldas.Columns.Count <= colIndex)
                        {
                            int colNumber = DGVCeldas.Columns.Count;
                            string colName = ((char)('A' + colNumber % 26)).ToString();
                            if (colNumber >= 26)
                                colName += ((char)('A' + (colNumber / 26) - 1)).ToString();
                            DGVCeldas.Columns.Add(colName, colName);
                        }

                        DGVCeldas[colIndex, rowIndex].Value = cells[j];
                    }
                }
                return;
            }
        }

        public void Movimiento(KeyEventArgs e)
        {
            int col = DGVCeldas.CurrentCell.ColumnIndex;
            int row = DGVCeldas.CurrentCell.RowIndex;

            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;

                if (row == DGVCeldas.Rows.Count - 1)
                {
                    DGVCeldas.Rows.Add();
                }

                if (row < DGVCeldas.Rows.Count - 1)
                {
                    DGVCeldas.CurrentCell = DGVCeldas[col, row + 1];
                }
            }

            else if (e.KeyCode == Keys.Tab || e.KeyCode == Keys.Right)
            {
                e.SuppressKeyPress = true;

                if (col == DGVCeldas.Columns.Count - 1)
                {
                    int i = DGVCeldas.Columns.Count;
                    string columnanomb = ((char)('A' + i % 26)).ToString();
                    if (i >= 26) columnanomb += ((char)('A' + (i / 26) - 1)).ToString();

                    DGVCeldas.Columns.Add(columnanomb, columnanomb);
                }

                if (col < DGVCeldas.Columns.Count - 1)
                {
                    DGVCeldas.CurrentCell = DGVCeldas[col + 1, row];
                }
            }
        }
    }
}
