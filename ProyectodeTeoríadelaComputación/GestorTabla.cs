using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProyectodeTeoríadelaComputación
{
    public class CopiaCelda
    {
        public object Valor { get; set; }
        public Font Fuente { get; set; }
        public Color ColorTexto { get; set; }
        public Color ColorFondo { get; set; }
        public DataGridViewContentAlignment Alineacion { get; set; }
    }

    public class GestorTabla
    {
        private DataGridView DGVCeldas;

        private List<List<CopiaCelda>> _copiaInterna = new List<List<CopiaCelda>>();

        public GestorTabla(DataGridView DGVCeldas)
        {
            this.DGVCeldas = DGVCeldas;
        }

        public void CopiarPegar(KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                Copiar();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                Pegar();
                e.Handled = true;
            }
        }

        private void Copiar()
        {
            _copiaInterna.Clear();

            var celdasOrdenadas = DGVCeldas.SelectedCells.Cast<DataGridViewCell>()
                .OrderBy(c => c.RowIndex)
                .ThenBy(c => c.ColumnIndex)
                .ToList();

            if (!celdasOrdenadas.Any()) return;

            int minRow = celdasOrdenadas.Min(c => c.RowIndex);
            int maxRow = celdasOrdenadas.Max(c => c.RowIndex);
            int minCol = celdasOrdenadas.Min(c => c.ColumnIndex);
            int maxCol = celdasOrdenadas.Max(c => c.ColumnIndex);

            for (int i = minRow; i <= maxRow; i++)
            {
                var fila = new List<CopiaCelda>();
                for (int j = minCol; j <= maxCol; j++)
                {
                    var cell = DGVCeldas[j, i];
                    var copia = new CopiaCelda
                    {
                        Valor = cell.Value,
                        Fuente = cell.Style.Font ?? DGVCeldas.DefaultCellStyle.Font ?? DGVCeldas.Font,
                        ColorTexto = cell.Style.ForeColor.IsEmpty ? DGVCeldas.DefaultCellStyle.ForeColor : cell.Style.ForeColor,
                        ColorFondo = cell.Style.BackColor.IsEmpty ? DGVCeldas.DefaultCellStyle.BackColor : cell.Style.BackColor,
                        Alineacion = cell.Style.Alignment == DataGridViewContentAlignment.NotSet ? DGVCeldas.DefaultCellStyle.Alignment : cell.Style.Alignment
                    };
                    fila.Add(copia);
                }
                _copiaInterna.Add(fila);
            }

            StringBuilder sb = new StringBuilder();
            foreach (var fila in _copiaInterna)
            {
                sb.AppendLine(string.Join("\t", fila.Select(c => c.Valor?.ToString() ?? "")));
            }
            Clipboard.SetText(sb.ToString().TrimEnd());
        }

        private void Pegar()
        {
            if (_copiaInterna == null || !_copiaInterna.Any()) return;

            int startRow = DGVCeldas.CurrentCell.RowIndex;
            int startCol = DGVCeldas.CurrentCell.ColumnIndex;

            for (int i = 0; i < _copiaInterna.Count; i++)
            {
                for (int j = 0; j < _copiaInterna[i].Count; j++)
                {
                    int rowIndex = startRow + i;
                    int colIndex = startCol + j;

                    while (DGVCeldas.Rows.Count <= rowIndex)
                        DGVCeldas.Rows.Add();

                    while (DGVCeldas.Columns.Count <= colIndex)
                    {
                        int colNumber = DGVCeldas.Columns.Count;
                        string colName = ReferenciasCeldas.ObtenerCeldaActual(colNumber);
                        DGVCeldas.Columns.Add(colName, colName);
                    }

                    var destCell = DGVCeldas[colIndex, rowIndex];
                    var copia = _copiaInterna[i][j];

                    destCell.Value = copia.Valor;
                    destCell.Style.Font = copia.Fuente;
                    destCell.Style.ForeColor = copia.ColorTexto;
                    destCell.Style.BackColor = copia.ColorFondo;
                    destCell.Style.Alignment = copia.Alineacion;
                }
            }
        }

        public Action ActualizarEncabezados;

        public void Movimiento(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                e.Handled = true;

                int col = DGVCeldas.CurrentCell.ColumnIndex;
                int row = DGVCeldas.CurrentCell.RowIndex;

                if (row == DGVCeldas.Rows.Count - 1)
                {
                    DGVCeldas.Rows.Add();

                    if (DGVCeldas.Rows.Count > 1)
                    {
                        int alto = DGVCeldas.Rows[row].Height;
                        DGVCeldas.Rows[DGVCeldas.Rows.Count - 1].Height = alto;
                    }

                    ActualizarEncabezados?.Invoke();
                }

                if (row < DGVCeldas.Rows.Count - 1)
                {
                    DGVCeldas.CurrentCell = DGVCeldas[col, row + 1];
                }
            }
            else if (e.KeyCode == Keys.Tab || e.KeyCode == Keys.Right)
            {
                e.SuppressKeyPress = true;

                int col = DGVCeldas.CurrentCell.ColumnIndex;
                int row = DGVCeldas.CurrentCell.RowIndex;

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

        private void Ordenar(bool ascendente)
        {
            if (DGVCeldas.SelectedCells.Count == 0) return;

            var selectedCells = DGVCeldas.SelectedCells.Cast<DataGridViewCell>();
            var colIndex = selectedCells.First().ColumnIndex;
            var rows = selectedCells.Select(cell => cell.RowIndex).Distinct().OrderBy(r => r).ToList();

            var lista = rows.Select(rowIndex => new
            {
                RowIndex = rowIndex,
                Value = DGVCeldas[colIndex, rowIndex].Value?.ToString() ?? ""
            }).ToList();

            bool allNumeric = lista.All(x => double.TryParse(x.Value, out _));

            if (allNumeric)
            {
                if (ascendente)
                    lista = lista.OrderBy(x => double.Parse(x.Value)).ToList();
                else
                    lista = lista.OrderByDescending(x => double.Parse(x.Value)).ToList();
            }
            else
            {
                if (ascendente)
                    lista = lista.OrderBy(x => x.Value).ToList();
                else
                    lista = lista.OrderByDescending(x => x.Value).ToList();
            }

            var sortedValues = lista.Select(x => x.Value).ToList();

            for (int i = 0; i < rows.Count; i++)
            {
                DGVCeldas[colIndex, rows[i]].Value = sortedValues[i];
            }
        }

        public void OrdenarAscendente()
        {
            Ordenar(true);
        }

        public void OrdenarDescendente()
        {
            Ordenar(false);
        }

        public void ConcatenarTexto()
        {
            if (DGVCeldas.SelectedCells.Count == 0)
            {
                MessageBox.Show("Seleccione celdas para concatenar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selectedCells = DGVCeldas.SelectedCells.Cast<DataGridViewCell>().ToList();

            var nonEmptyCells = selectedCells.Where(c => c.Value != null && !string.IsNullOrWhiteSpace(c.Value.ToString())).ToList();
            var emptyCells = selectedCells.Where(c => c.Value == null || string.IsNullOrWhiteSpace(c.Value.ToString())).ToList();

            if (nonEmptyCells.Count == 0)
            {
                MessageBox.Show("No hay celdas con contenido para concatenar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (emptyCells.Count == 0)
            {
                MessageBox.Show("Seleccione al menos una celda vacía para colocar el texto concatenado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            List<string> textos = new List<string>();

            var orderedNonEmpty = nonEmptyCells.OrderBy(c => c.RowIndex).ThenBy(c => c.ColumnIndex).ToList();

            foreach (var cell in orderedNonEmpty)
            {
                textos.Add(cell.Value.ToString());
            }

            string resultado = string.Join(" ", textos);

            var destino = emptyCells.OrderBy(c => c.RowIndex).ThenBy(c => c.ColumnIndex).First();
            destino.Value = resultado;
        }
    }
}
