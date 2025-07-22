using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProyectodeTeoríadelaComputación
{
    public partial class Libro1 : Form
    {
        public Libro1()
        {
            InitializeComponent();
        }

        private void Libro1_Load(object sender, EventArgs e)
        {
            DGVCeldas.Rows.Clear();
            DGVCeldas.Rows.Add(32);

            foreach (FontFamily font in System.Drawing.FontFamily.Families)
            {
                cmbFuentes.Items.Add(font.Name);
            }

            cmbFuentes.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbFuentes.SelectedIndexChanged += cmbFuentes_SelectedIndexChanged;

            for (int i = 8; i <= 72; i += 2)
            {
                cmbTamaño.Items.Add(i);
            }

            cmbTamaño.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbTamaño.SelectedIndexChanged += cmbTamaño_SelectedIndexChanged;
        }

        private void DGVCeldas_KeyDown(object sender, KeyEventArgs e)
        {
            CopiarPegar(e);
            Movimiento(e);
        }

        private void CopiarPegar(KeyEventArgs e)
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

        private void Movimiento(KeyEventArgs e)
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

        private void DGVCeldas_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            string rowNumber = (e.RowIndex + 1).ToString();

            var size = e.Graphics.MeasureString(rowNumber, DGVCeldas.Font);
            int leftMargin = 5;
            int topMargin = (e.RowBounds.Height - (int)size.Height) / 2;

            e.Graphics.DrawString(
                rowNumber,
                DGVCeldas.Font,
                SystemBrushes.ControlText,
                e.RowBounds.Left + leftMargin,
                e.RowBounds.Top + topMargin
            );
        }

        private void btnNegrita_Click(object sender, EventArgs e)
        {
            if (DGVCeldas.SelectedCells.Count == 0) return;

            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                Font currentFont = cell.Style.Font ?? DGVCeldas.DefaultCellStyle.Font ?? DGVCeldas.Font;

                FontStyle newStyle = currentFont.Style ^ FontStyle.Bold;

                cell.Style.Font = new Font(currentFont, newStyle);
            }
        }

        private void btnCursiva_Click(object sender, EventArgs e)
        {
            if (DGVCeldas.SelectedCells.Count == 0) return;

            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                Font currentFont = cell.Style.Font ?? DGVCeldas.DefaultCellStyle.Font ?? DGVCeldas.Font;

                FontStyle newStyle = currentFont.Style ^ FontStyle.Italic;

                cell.Style.Font = new Font(currentFont, newStyle);
            }
        }

        private void btnSubrayado_Click(object sender, EventArgs e)
        {
            if (DGVCeldas.SelectedCells.Count == 0) return;

            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                Font currentFont = cell.Style.Font ?? DGVCeldas.DefaultCellStyle.Font ?? DGVCeldas.Font;
                FontStyle newStyle = currentFont.Style ^ FontStyle.Underline;

                cell.Style.Font = new Font(currentFont.FontFamily, currentFont.Size, newStyle);
            }
        }

        private void cmbFuentes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DGVCeldas.SelectedCells.Count == 0) return;

            string selectedFont = cmbFuentes.SelectedItem.ToString();

            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                Font currentFont = cell.Style.Font ?? DGVCeldas.DefaultCellStyle.Font ?? DGVCeldas.Font;
                cell.Style.Font = new Font(selectedFont, currentFont.Size, currentFont.Style);
            }
        }

        private void cmbTamaño_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DGVCeldas.SelectedCells.Count == 0) return;

            float size = float.Parse(cmbTamaño.SelectedItem.ToString());

            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                Font currentFont = cell.Style.Font ?? DGVCeldas.DefaultCellStyle.Font ?? DGVCeldas.Font;
                cell.Style.Font = new Font(currentFont.FontFamily, size, currentFont.Style);
            }
        }

        private void DGVCeldas_CurrentCellChanged(object sender, EventArgs e)
        {
            if (DGVCeldas.CurrentCell == null) return;

            int colIndex = DGVCeldas.CurrentCell.ColumnIndex;
            int rowIndex = DGVCeldas.CurrentCell.RowIndex;

            string colLetter = ObtenerCeldaActual(colIndex);
            txtCeldaActiva.Text = $"{colLetter}{rowIndex + 1}";
        }

        private string ObtenerCeldaActual(int colIndex)
        {
            int dividend = colIndex + 1;
            string columnName = String.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        private void txtCeldaActiva_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;

                string input = txtCeldaActiva.Text.Trim().ToUpper();

                int i = 0;
                while (i < input.Length && Char.IsLetter(input[i]))
                    i++;

                if (i == 0 || i >= input.Length) return;

                string colPart = input.Substring(0, i);
                string rowPart = input.Substring(i);

                int colIndex = ObtenerNumColumna(colPart);
                int rowIndex;
                if (!int.TryParse(rowPart, out rowIndex)) return;

                rowIndex -= 1;

                int maxColIndex = ObtenerNumColumna("AA");
                int maxRowIndex = 500 - 1;

                if (colIndex < 0 || colIndex > maxColIndex)
                {
                    MessageBox.Show($"La columna excede el límite máximo de AA.");
                    return;
                }

                if (rowIndex < 0 || rowIndex > maxRowIndex)
                {
                    MessageBox.Show($"La fila excede el límite máximo de 500.");
                    return;
                }

                if (DGVCeldas.Columns.Count <= colIndex)
                {
                    int columnsToAdd = colIndex - DGVCeldas.Columns.Count + 1;
                    for (int c = 0; c < columnsToAdd; c++)
                    {
                        string colName = ObtenerCeldaActual(DGVCeldas.Columns.Count);
                        DGVCeldas.Columns.Add(colName, colName);
                    }
                }

                if (DGVCeldas.Rows.Count <= rowIndex)
                {
                    int rowsToAdd = rowIndex - DGVCeldas.Rows.Count + 1;
                    DGVCeldas.Rows.Add(rowsToAdd);
                }

                DGVCeldas.CurrentCell = DGVCeldas[colIndex, rowIndex];
            }
        }

        private int ObtenerNumColumna(string colLetters)
        {
            int col = 0;
            for (int i = 0; i < colLetters.Length; i++)
            {
                col *= 26;
                col += (colLetters[i] - 'A' + 1);
            }
            return col - 1;
        }

        private void btnTXIzquierda_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                cell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
            }
        }

        private void btnTXCentro_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private void btnTXDerecha_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                cell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
        }

        private void btnTXArriba_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                var currentAlignment = cell.Style.Alignment;

                var horizontal = ObtenerAlineación(currentAlignment);
                cell.Style.Alignment = CombinarAlineación(horizontal, "Arriba");
            }
        }

        private void btnTXCentroV_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                var currentAlignment = cell.Style.Alignment;

                var horizontal = ObtenerAlineación(currentAlignment);

                cell.Style.Alignment = CombinarAlineación(horizontal, "Centro");
            }
        }

        private void btnTXAbajo_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
            {
                var currentAlignment = cell.Style.Alignment;

                var horizontal = ObtenerAlineación(currentAlignment);

                cell.Style.Alignment = CombinarAlineación(horizontal, "Abajo");
            }
        }

        private string ObtenerAlineación(DataGridViewContentAlignment alignment)
        {
            switch (alignment)
            {
                case DataGridViewContentAlignment.TopLeft:
                case DataGridViewContentAlignment.MiddleLeft:
                case DataGridViewContentAlignment.BottomLeft:
                    return "Izquierda";
                case DataGridViewContentAlignment.TopCenter:
                case DataGridViewContentAlignment.MiddleCenter:
                case DataGridViewContentAlignment.BottomCenter:
                    return "Medio";
                case DataGridViewContentAlignment.TopRight:
                case DataGridViewContentAlignment.MiddleRight:
                case DataGridViewContentAlignment.BottomRight:
                    return "Derecha";
                default:
                    return "Izquierda";
            }
        }

        private DataGridViewContentAlignment CombinarAlineación(string horizontal, string vertical)
        {
            switch (vertical + horizontal)
            {
                case "ArribaIzquierda": return DataGridViewContentAlignment.TopLeft;
                case "ArribaMedio": return DataGridViewContentAlignment.TopCenter;
                case "ArribaDerecha": return DataGridViewContentAlignment.TopRight;
                case "CentroIzquierda": return DataGridViewContentAlignment.MiddleLeft;
                case "CentroMedio": return DataGridViewContentAlignment.MiddleCenter;
                case "CentroDerecha": return DataGridViewContentAlignment.MiddleRight;
                case "AbajoIzquierda": return DataGridViewContentAlignment.BottomLeft;
                case "AbajoMedio": return DataGridViewContentAlignment.BottomCenter;
                case "AbajoDerecha": return DataGridViewContentAlignment.BottomRight;
                default: return DataGridViewContentAlignment.MiddleLeft;
            }
        }

        private void btnRellenoColor_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    Color selectedColor = colorDialog.Color;

                    foreach (DataGridViewCell cell in DGVCeldas.SelectedCells)
                    {
                        cell.Style.BackColor = selectedColor;
                    }
                }
            }
        }
    }
}
