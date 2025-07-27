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
        private GestorTabla gesttabla;
        private FormatoCeldas formatcelda;

        public Libro1()
        {
            InitializeComponent();
            gesttabla = new GestorTabla(DGVCeldas);
            formatcelda = new FormatoCeldas(DGVCeldas);
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
            gesttabla.CopiarPegar(e);
            gesttabla.Movimiento(e);
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
            formatcelda.AplicarNegrita();
        }

        private void btnCursiva_Click(object sender, EventArgs e)
        {
            formatcelda.AplicarCursiva();
        }

        private void btnSubrayado_Click(object sender, EventArgs e)
        {
            formatcelda.AplicarSubrayado();
        }

        private void cmbFuentes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbFuentes.SelectedItem != null)
            {
                formatcelda.CambiarFuente(cmbFuentes.SelectedItem.ToString());
            }
        }

        private void cmbTamaño_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbTamaño.SelectedItem != null)
            {
                if (float.TryParse(cmbTamaño.SelectedItem.ToString(), out float size))
                {
                    formatcelda.CambiarTamaño(size);
                }
            }
        }

        private void DGVCeldas_CurrentCellChanged(object sender, EventArgs e)
        {
            if (DGVCeldas.CurrentCell == null) return;

            int colIndex = DGVCeldas.CurrentCell.ColumnIndex;
            int rowIndex = DGVCeldas.CurrentCell.RowIndex;

            string colLetter = ReferenciasCeldas.ObtenerCeldaActual(colIndex);
            txtCeldaActiva.Text = $"{colLetter}{rowIndex + 1}";
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

                int colIndex = ReferenciasCeldas.ObtenerNumColumna(colPart);
                int rowIndex;
                if (!int.TryParse(rowPart, out rowIndex)) return;

                rowIndex -= 1;

                int maxColIndex = ReferenciasCeldas.ObtenerNumColumna("AA");
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
                        string colName = ReferenciasCeldas.ObtenerCeldaActual(DGVCeldas.Columns.Count);
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

        private string ObtenerAlineación(DataGridViewContentAlignment alineación)
        {
            switch (alineación)
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
                    formatcelda.AplicarRellenoColor(selectedColor);
                }
            }
        }
    }
}
