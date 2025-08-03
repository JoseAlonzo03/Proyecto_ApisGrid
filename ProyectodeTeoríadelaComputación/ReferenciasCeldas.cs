using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProyectodeTeoríadelaComputación
{
    internal class ReferenciasCeldas
    {
        public static string ObtenerCeldaActual(int index)
        {
            const string letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string resultado = "";

            while (index >= 0)
            {
                resultado = letras[index % 26] + resultado;
                index = index / 26 - 1;
            }

            return resultado;
        }

        public static int ObtenerNumColumna(string referencia)
        {
            int result = 0;
            for (int i = 0; i < referencia.Length; i++)
            {
                result *= 26;
                result += referencia[i] - 'A' + 1;
            }
            return result - 1;
        }

        public static void MostrarEncabezados(DataGridView dgv, bool mostrar)
        {
            dgv.ColumnHeadersVisible = mostrar;
            dgv.RowHeadersVisible = mostrar;
        }

        public static void MostrarLineasDivision(DataGridView dgv, bool mostrar)
        {
            dgv.CellBorderStyle = mostrar ? DataGridViewCellBorderStyle.Single : DataGridViewCellBorderStyle.None;
        }

        public static void MostrarBarraFormulas(TextBox barra, DataGridView dgv, bool mostrar)
        {
            barra.Visible = mostrar;

            if (mostrar)
            {
                dgv.Location = new Point(dgv.Location.X, barra.Bottom + 5);
                dgv.Height -= barra.Height + 5;
            }
            else
            {
                dgv.Location = new Point(dgv.Location.X, barra.Location.Y);
                dgv.Height += barra.Height + 5;
            }
        }

        public static void ValidarYAplicarZoom(TextBox txtZoom, DataGridView dgv, TrackBar trkZoom)
        {
            string input = txtZoom.Text.Trim().Replace("%", "");

            if (int.TryParse(input, out int zoom))
            {
                if (zoom < trkZoom.Minimum)
                {
                    MessageBox.Show($"El zoom mínimo es {trkZoom.Minimum}%.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    zoom = trkZoom.Minimum;
                }
                else if (zoom > trkZoom.Maximum)
                {
                    MessageBox.Show($"El zoom máximo es {trkZoom.Maximum}%.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    zoom = trkZoom.Maximum;
                }

                txtZoom.Text = $"{zoom}%";
                trkZoom.Value = zoom;

                AplicarZoom(dgv, zoom);
            }
            else
            {
                MessageBox.Show("Ingrese un valor numérico válido para el zoom.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                zoom = 100;
                txtZoom.Text = $"{zoom}%";
                trkZoom.Value = zoom;

                AplicarZoom(dgv, zoom);
            }
        }


        public static void AplicarZoom(DataGridView dgv, int porcentaje)
        {
            float zoom = porcentaje / 100.0f;

            foreach (DataGridViewColumn col in dgv.Columns)
            {
                col.DefaultCellStyle.Font = new Font(dgv.Font.FontFamily, 11 * zoom);
            }

            dgv.RowHeadersDefaultCellStyle.Font = new Font(dgv.Font.FontFamily, 11 * zoom);
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font(dgv.Font.FontFamily, 11 * zoom);

            int anchoBase = 50;
            dgv.RowHeadersWidth = (int)(anchoBase * zoom);

            int altoBaseFila = 22;
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.Height = (int)(altoBaseFila * zoom);
            }

            int altoBaseEncabezado = 25;
            dgv.ColumnHeadersHeight = (int)(altoBaseEncabezado * zoom);
        }

        public static void ProcesarBarraFormulasEnter(DataGridView dgv, TextBox txtBarraFormulas, FormulasOperaciones formOpera)
        {
            if (dgv.CurrentCell == null) return;

            int colIndex = dgv.CurrentCell.ColumnIndex;
            int rowIndex = dgv.CurrentCell.RowIndex;

            string input = txtBarraFormulas.Text.Trim();

            dgv[colIndex, rowIndex].Value = input;

            if (input.StartsWith("="))
            {
                formOpera.EvaluarYAsignar(colIndex, rowIndex, input);
            }
        }

        private static Dictionary<(int col, int row), string> formulasGuardadas = new Dictionary<(int col, int row), string>();
        private static Dictionary<(int col, int row), List<(int col, int row)>> dependencias = new Dictionary<(int col, int row), List<(int col, int row)>>();

        public static string ObtenerFormula((int col, int row) celda)
        {
            if (formulasGuardadas.TryGetValue(celda, out string formula))
                return formula;
            return null;
        }

        public static void GuardarFormula((int col, int row) celda, string formula)
        {
            formulasGuardadas[celda] = formula;
            ActualizarDependencias(celda, formula);
        }

        public static void RemoverFormula((int col, int row) celda)
        {
            if (formulasGuardadas.ContainsKey(celda))
                formulasGuardadas.Remove(celda);

            foreach (var entry in dependencias)
            {
                entry.Value.Remove(celda);
            }
        }

        public static List<(int col, int row)> ExtraerReferencias(string formula)
        {
            var refs = new List<(int col, int row)>();
            if (string.IsNullOrEmpty(formula)) return refs;

            var regex = new Regex(@"([A-Z]+)(\d+)", RegexOptions.IgnoreCase);

            foreach (Match match in regex.Matches(formula))
            {
                string colLetra = match.Groups[1].Value.ToUpper();
                if (!int.TryParse(match.Groups[2].Value, out int fila)) continue;

                int col = ObtenerNumColumna(colLetra);
                int row = fila - 1;
                refs.Add((col, row));
            }

            return refs;
        }

        private static void ActualizarDependencias((int col, int row) celda, string formula)
        {
            foreach (var entry in dependencias)
            {
                entry.Value.Remove(celda);
            }

            var referencias = ExtraerReferencias(formula);

            foreach (var referencia in referencias)
            {
                if (!dependencias.ContainsKey(referencia))
                    dependencias[referencia] = new List<(int col, int row)>();

                if (!dependencias[referencia].Contains(celda))
                    dependencias[referencia].Add(celda);
            }
        }

        public static List<(int col, int row)> ObtenerDependientes((int col, int row) celda)
        {
            if (dependencias.ContainsKey(celda))
                return dependencias[celda];
            return new List<(int col, int row)>();
        }

        public static void RecalcularDependientes(DataGridView dgv, FormulasOperaciones formOpera, (int col, int row) celda)
        {
            var dependientes = ObtenerDependientes(celda);
            foreach (var dep in dependientes)
            {
                if (formulasGuardadas.TryGetValue(dep, out string formula))
                {
                    formOpera.EvaluarYAsignar(dep.col, dep.row, formula);
                    dgv[dep.col, dep.row].Value = dgv[dep.col, dep.row].Value;
                    RecalcularDependientes(dgv, formOpera, dep);
                }
            }
        }
    }
}
