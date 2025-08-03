using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProyectodeTeoríadelaComputación
{
    public partial class Libro1 : Form
    {
        private GestorTabla gestTabla;
        private FormatoCeldas formatCelda;
        private FormulasOperaciones formOpera;
        public DataGridView TablaCeldas => DGVCeldas;

        private bool isOverCornerHeader = false;

        public Libro1()
        {
            InitializeComponent();
            gestTabla = new GestorTabla(DGVCeldas);
            gestTabla.ActualizarEncabezados = ActualizarNumerosDeFila;
            formatCelda = new FormatoCeldas(DGVCeldas);
            formOpera = new FormulasOperaciones(DGVCeldas);

            DGVCeldas.RowsAdded += DGVCeldas_RowsAdded;
            DGVCeldas.CellMouseEnter += DGVCeldas_CellMouseEnter;
            DGVCeldas.CellMouseLeave += DGVCeldas_CellMouseLeave;
            DGVCeldas.CellDoubleClick += DGVCeldas_CellDoubleClick;


            DGVCeldas.MouseMove += DGVCeldas_MouseMove;
            DGVCeldas.MouseLeave += DGVCeldas_MouseLeave;

            txtBarraFormulas.KeyDown += txtBarraFormulas_KeyDown;
        }

        private void Libro1_Load(object sender, EventArgs e)
        {
            DGVCeldas.Columns.Clear();

            for (int i = 0; i < 26; i++)
            {
                string colName = ReferenciasCeldas.ObtenerCeldaActual(i);
                DGVCeldas.Columns.Add(colName, colName);
            }

            DGVCeldas.Columns[0].Width = 80;
            for (int i = 1; i < DGVCeldas.Columns.Count; i++)
            {
                DGVCeldas.Columns[i].Width = 80;
            }

            DGVCeldas.Rows.Clear();
            DGVCeldas.Rows.Add(32);
            ActualizarNumerosDeFila();

            cmbFuentes.Items.Clear();
            foreach (FontFamily font in System.Drawing.FontFamily.Families)
            {
                cmbFuentes.Items.Add(font.Name);
            }
            cmbFuentes.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbFuentes.SelectedIndexChanged += cmbFuentes_SelectedIndexChanged;

            cmbTamaño.Items.Clear();
            for (int i = 8; i <= 72; i += 2)
            {
                cmbTamaño.Items.Add(i);
            }
            cmbTamaño.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbTamaño.SelectedIndexChanged += cmbTamaño_SelectedIndexChanged;
        }

        private void DGVCeldas_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            ActualizarNumerosDeFila();
        }

        private void ActualizarNumerosDeFila()
        {
            for (int i = 0; i < DGVCeldas.Rows.Count; i++)
            {
                DGVCeldas.Rows[i].HeaderCell.Value = (i + 1).ToString();
            }
        }

        private void DGVCeldas_KeyDown(object sender, KeyEventArgs e)
        {
            gestTabla.Movimiento(e);
            gestTabla.CopiarPegar(e);
        }

        private void btnNegrita_Click(object sender, EventArgs e)
        {
            formatCelda.AplicarNegrita();
        }

        private void btnCursiva_Click(object sender, EventArgs e)
        {
            formatCelda.AplicarCursiva();
        }

        private void btnSubrayado_Click(object sender, EventArgs e)
        {
            formatCelda.AplicarSubrayado();
        }

        private void cmbFuentes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbFuentes.SelectedItem != null)
            {
                formatCelda.CambiarFuente(cmbFuentes.SelectedItem.ToString());
            }
        }

        private void cmbTamaño_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbTamaño.SelectedItem != null)
            {
                if (float.TryParse(cmbTamaño.SelectedItem.ToString(), out float size))
                {
                    formatCelda.CambiarTamaño(size);
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

            var formula = ReferenciasCeldas.ObtenerFormula((colIndex, rowIndex));

            if (!string.IsNullOrEmpty(formula))
            {
                txtBarraFormulas.Text = formula;
            }
            else
            {
                var valor = DGVCeldas.CurrentCell.Value;
                txtBarraFormulas.Text = valor?.ToString() ?? "";
            }
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
                        DGVCeldas.Columns[DGVCeldas.Columns.Count - 1].Width = 60;
                        DGVCeldas.Columns[DGVCeldas.Columns.Count - 1].Resizable = DataGridViewTriState.True;
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
                    formatCelda.AplicarRellenoColor(selectedColor);
                }
            }
        }
        private void btnColorTexto_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    Color selectedColor = colorDialog.Color;
                    formatCelda.AplicarColorTexto(selectedColor);
                }
            }
        }

        private void btnOrdenAsc_Click(object sender, EventArgs e)
        {
           gestTabla.OrdenarAscendente();
        }

        private void btnOrdenDes_Click(object sender, EventArgs e)
        {
           gestTabla.OrdenarDescendente();
        }

        private void DGVCeldas_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex >= 0)
            {
                DGVCeldas.Columns[e.ColumnIndex].HeaderCell.Style.BackColor = Color.LightBlue;
            }
            else if (e.ColumnIndex == -1 && e.RowIndex >= 0)
            {
                DGVCeldas.Rows[e.RowIndex].HeaderCell.Style.BackColor = Color.LightBlue;
            }
        }

        private void DGVCeldas_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex >= 0)
            {
                DGVCeldas.Columns[e.ColumnIndex].HeaderCell.Style.BackColor = SystemColors.Control;
            }
            else if (e.ColumnIndex == -1 && e.RowIndex >= 0)
            {
                DGVCeldas.Rows[e.RowIndex].HeaderCell.Style.BackColor = SystemColors.Control;
            }
        }

        private void DGVCeldas_MouseMove(object sender, MouseEventArgs e)
        {
            Rectangle cornerRect = new Rectangle(
                0, 0,
                DGVCeldas.RowHeadersWidth,
                DGVCeldas.ColumnHeadersHeight
            );

            Point clientPoint = new Point(e.X, e.Y);

            if (cornerRect.Contains(clientPoint))
            {
                if (!isOverCornerHeader)
                {
                    isOverCornerHeader = true;
                    DGVCeldas.TopLeftHeaderCell.Style.BackColor = Color.LightBlue;
                }
            }
            else
            {
                if (isOverCornerHeader)
                {
                    isOverCornerHeader = false;
                    DGVCeldas.TopLeftHeaderCell.Style.BackColor = SystemColors.Control;
                }
            }
        }

        private void DGVCeldas_MouseLeave(object sender, EventArgs e)
        {
            if (isOverCornerHeader)
            {
                isOverCornerHeader = false;
                DGVCeldas.TopLeftHeaderCell.Style.BackColor = SystemColors.Control;
            }
        }

        private void btnAjustarTodo_Click(object sender, EventArgs e)
        {
            formatCelda.AutoAjustarTodo();
        }

        private void btnLimpiarCeldas_Click(object sender, EventArgs e)
        {
            formatCelda.LimpiarCeldas(txtBarraFormulas);
        }

        private void DGVCeldas_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            var cell = DGVCeldas[e.ColumnIndex, e.RowIndex];
            var key = (e.ColumnIndex, e.RowIndex);
            string texto = cell.Value?.ToString() ?? "";

            if (texto.StartsWith("="))
            {
                ReferenciasCeldas.GuardarFormula(key, texto);
                formOpera.EvaluarYAsignar(e.ColumnIndex, e.RowIndex, texto);
                cell.Value = DGVCeldas[e.ColumnIndex, e.RowIndex].Value;
            }
            else
            {
                ReferenciasCeldas.RemoverFormula(key);
            }

            ReferenciasCeldas.RecalcularDependientes(DGVCeldas, formOpera, key);

            txtBarraFormulas.Text = cell.Value?.ToString() ?? "";
        }


        private void btnSuma_Click(object sender, EventArgs e)
        {
            formOpera.EvaluarOperacionSeleccionada(DGVCeldas, "SUMA");
        }

        private void btnResta_Click(object sender, EventArgs e)
        {
            formOpera.EvaluarOperacionSeleccionada(DGVCeldas, "RESTA");
        }

        private void btnMultiplicar_Click(object sender, EventArgs e)
        {
            formOpera.EvaluarOperacionSeleccionada(DGVCeldas, "MULTI");
        }

        private void btnDividir_Click(object sender, EventArgs e)
        {
            formOpera.EvaluarOperacionSeleccionada(DGVCeldas, "DIVIDIR");
        }

        private void btnConcatenar_Click(object sender, EventArgs e)
        {
            gestTabla.ConcatenarTexto();
        }

        private void btnMínimo_Click(object sender, EventArgs e)
        {
            formOpera.EvaluarOperacionSeleccionada(DGVCeldas, "MIN");
        }

        private void btnMáximo_Click(object sender, EventArgs e)
        {
            formOpera.EvaluarOperacionSeleccionada(DGVCeldas, "MAX");
        }

        private void btnPromedio_Click(object sender, EventArgs e)
        {
            formOpera.EvaluarOperacionSeleccionada(DGVCeldas, "PROM");
        }

        private void btnContarValor_Click(object sender, EventArgs e)
        {
            formOpera.EvaluarOperacionSeleccionada(DGVCeldas, "CONTVAL");
        }

        private void btnPotencia_Click(object sender, EventArgs e)
        {
            formOpera.EvaluarOperacionSeleccionada(DGVCeldas, "POWER");
        }

        private void btnRaíz_Click(object sender, EventArgs e)
        {
            formOpera.EvaluarOperacionSeleccionada(DGVCeldas, "SQRT");
        }

        private void btnValAbsoluto_Click(object sender, EventArgs e)
        {
            formOpera.EvaluarOperacionSeleccionada(DGVCeldas, "ABS");
        }

        private void btnModa_Click(object sender, EventArgs e)
        {
            formOpera.EvaluarOperacionSeleccionada(DGVCeldas, "MODA");
        }

        private void btnFecha_Click(object sender, EventArgs e)
        {
            var celdaResultado = DGVCeldas.SelectedCells.Cast<DataGridViewCell>()
                            .FirstOrDefault(c => c.Value == null || string.IsNullOrWhiteSpace(c.Value.ToString()));

            if (celdaResultado == null)
            {
                MessageBox.Show("Debe seleccionar una celda vacía para mostrar la fecha.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DGVCeldas[celdaResultado.ColumnIndex, celdaResultado.RowIndex].Value = DateTime.Now.ToString("dd/MM/yyyy");
        }
        
        private void btnHora_Click(object sender, EventArgs e)
        {
            var celdaResultado = DGVCeldas.SelectedCells.Cast<DataGridViewCell>()
                           .FirstOrDefault(c => c.Value == null || string.IsNullOrWhiteSpace(c.Value.ToString()));

            if (celdaResultado == null)
            {
                MessageBox.Show("Debe seleccionar una celda vacía para mostrar la hora.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DGVCeldas[celdaResultado.ColumnIndex, celdaResultado.RowIndex].Value = DateTime.Now.ToString("HH:mm:ss");
        }

        private void chkEncabezados_CheckedChanged(object sender, EventArgs e)
        {
            ReferenciasCeldas.MostrarEncabezados(DGVCeldas, chkEncabezados.Checked);
        }

        private void chkLíneaDiv_CheckedChanged(object sender, EventArgs e)
        {
            ReferenciasCeldas.MostrarLineasDivision(DGVCeldas, chkLíneaDiv.Checked);
        }

        private void chkBarraForm_CheckedChanged(object sender, EventArgs e)
        {
            ReferenciasCeldas.MostrarBarraFormulas(txtBarraFormulas, DGVCeldas, chkBarraForm.Checked);
        }

        private void txtBarraFormulas_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;

                if (DGVCeldas.CurrentCell != null)
                {
                    int colIndex = DGVCeldas.CurrentCell.ColumnIndex;
                    int rowIndex = DGVCeldas.CurrentCell.RowIndex;
                    var key = (colIndex, rowIndex);

                    string texto = txtBarraFormulas.Text;

                    DGVCeldas[colIndex, rowIndex].Value = texto;

                    if (texto.StartsWith("="))
                    {
                        ReferenciasCeldas.GuardarFormula(key, texto);
                        formOpera.EvaluarYAsignar(colIndex, rowIndex, texto);
                        ReferenciasCeldas.RecalcularDependientes(DGVCeldas, formOpera, key);
                    }
                    else
                    {
                        ReferenciasCeldas.RemoverFormula(key);
                    }
                }
            }
        }

        private void DGVCeldas_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var key = (e.ColumnIndex, e.RowIndex);
            var formula = ReferenciasCeldas.ObtenerFormula(key);

            if (!string.IsNullOrEmpty(formula))
            {
                DGVCeldas[e.ColumnIndex, e.RowIndex].Value = formula;
                txtBarraFormulas.Text = formula;
            }
            else
            {
                txtBarraFormulas.Text = DGVCeldas[e.ColumnIndex, e.RowIndex].Value?.ToString() ?? "";
            }

            DGVCeldas.BeginEdit(true);
        }

        private void TrkBZoom_Scroll(object sender, EventArgs e)
        {
            ReferenciasCeldas.AplicarZoom(DGVCeldas, TrkBZoom.Value);
            txtZoom.Text = $"{TrkBZoom.Value}%";
        }

        private void txtZoom_Leave(object sender, EventArgs e)
        {
            ReferenciasCeldas.ValidarYAplicarZoom(txtZoom, DGVCeldas, TrkBZoom);
        }

        private void txtZoom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                ReferenciasCeldas.ValidarYAplicarZoom(txtZoom, DGVCeldas, TrkBZoom);
            }
        }

        private void archivoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("¿Desea guardar los cambios?", "Guardar Cambios",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                var guardarPdf = MessageBox.Show("¿Desea guardar como archivo PDF?", "Guardar como PDF",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (guardarPdf == DialogResult.Yes)
                {
                    GuardarDatos.ExportarPDF(DGVCeldas);
                    MostrarPantallaInicio();
                }
                else
                {
                    var guardarExcel = MessageBox.Show("¿Desea guardar como archivo de Excel?", "Guardar como Excel",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (guardarExcel == DialogResult.Yes)
                    {
                        GuardarDatos.ExportarExcel(DGVCeldas);
                        MostrarPantallaInicio();
                    }
                    else
                    {
                        MostrarPantallaInicio();
                    }
                }
            }
            else if (result == DialogResult.No)
            {
                MostrarPantallaInicio();
            }
        }

        private void nuevoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool estaVacio = EstaTablaVacia();

            if (estaVacio)
            {
                var confirmar = MessageBox.Show("El libro actual está vacío. ¿Desea abrir un nuevo libro?",
                    "Confirmar Nuevo Libro", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (confirmar == DialogResult.Yes)
                {
                    AbrirNuevoLibro1();
                }
            }
            else
            {
                var resultado = MessageBox.Show("¿Desea guardar los cambios antes de abrir un nuevo libro?",
                    "Guardar Cambios", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (resultado == DialogResult.Yes)
                {
                    var guardarPdf = MessageBox.Show("¿Desea guardar como archivo PDF?", "Guardar como PDF",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (guardarPdf == DialogResult.Yes)
                    {
                        GuardarDatos.ExportarPDF(DGVCeldas);
                        AbrirNuevoLibro1();
                    }
                    else
                    {
                        var guardarExcel = MessageBox.Show("¿Desea guardar como archivo de Excel?", "Guardar como Excel",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (guardarExcel == DialogResult.Yes)
                        {
                            GuardarDatos.ExportarExcel(DGVCeldas);
                            AbrirNuevoLibro1();
                        }
                        else
                        {
                            AbrirNuevoLibro1();
                        }
                    }
                }
                else if (resultado == DialogResult.No)
                {
                    AbrirNuevoLibro1();
                }
            }
        }

        private void abrirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool estaVacio = EstaTablaVacia();

            if (!estaVacio)
            {
                var resultado = MessageBox.Show("¿Desea guardar los cambios antes de abrir un archivo?",
                    "Guardar Cambios", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (resultado == DialogResult.Yes)
                {
                    var guardarPdf = MessageBox.Show("¿Desea guardar como archivo PDF?", "Guardar como PDF",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (guardarPdf == DialogResult.Yes)
                    {
                        GuardarDatos.ExportarPDF(DGVCeldas);
                    }
                    else
                    {
                        var guardarExcel = MessageBox.Show("¿Desea guardar como archivo de Excel?", "Guardar como Excel",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (guardarExcel == DialogResult.Yes)
                        {
                            GuardarDatos.ExportarExcel(DGVCeldas);
                        }
                    }
                }
                else if (resultado == DialogResult.Cancel)
                {
                    return;
                }
            }

            using (OpenFileDialog ofd = new OpenFileDialog()
            {
                Filter = "Archivos Excel (*.xlsx)|*.xlsx",
                Title = "Abrir archivo Excel"
            })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    GuardarDatos.ImportarExcel(DGVCeldas, ofd.FileName);
                    ActualizarNumerosDeFila();
                }
            }
        }

        private void fórmulasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormFórmulas form = new FormFórmulas();
            form.Show(this);
        }

        private void MostrarPantallaInicio()
        {
            PantallaInicio inicio = new PantallaInicio();
            inicio.Show();
            this.Close();
        }

        private void AbrirNuevoLibro1()
        {
            Libro1 nuevoLibro = new Libro1();
            nuevoLibro.Show();
            this.Close();
        }

        public void CargarArchivoExcel(string rutaArchivo)
        {
            GuardarDatos.ImportarExcel(DGVCeldas, rutaArchivo);
            ActualizarNumerosDeFila();
        }

        private bool EstaTablaVacia()
        {
            if (DGVCeldas.Rows.Count == 0) return true;

            foreach (DataGridViewRow fila in DGVCeldas.Rows)
            {
                if (fila.IsNewRow) continue;

                foreach (DataGridViewCell celda in fila.Cells)
                {
                    if (celda.Value != null && !string.IsNullOrWhiteSpace(celda.Value.ToString()))
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GuardarDatos.ExportarPDF(DGVCeldas);
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GuardarDatos.ExportarExcel(DGVCeldas);
        }
    }
}