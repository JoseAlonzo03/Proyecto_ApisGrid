using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProyectodeTeoríadelaComputación
{
    public class FormulasOperaciones
    {
        private DataGridView DGVCeldas;

        public FormulasOperaciones(DataGridView dgv)
        {
            DGVCeldas = dgv;
        }

        public void EvaluarYAsignar(int colIndex, int rowIndex, string formula)
        {
            try
            {
                if (Regex.IsMatch(formula, @"^\s*=\s*FECHA\s*\(\s*\)\s*$", RegexOptions.IgnoreCase))
                {
                    DGVCeldas[colIndex, rowIndex].Value = DateTime.Now.ToString("dd/MM/yyyy");
                    return;
                }
                if (Regex.IsMatch(formula, @"^\s*=\s*HORA\s*\(\s*\)\s*$", RegexOptions.IgnoreCase))
                {
                    DGVCeldas[colIndex, rowIndex].Value = DateTime.Now.ToString("HH:mm:ss");
                    return;
                }

                string resultado = EvaluarFormula(formula);
                DGVCeldas[colIndex, rowIndex].Value = resultado;
            }
            catch (DivideByZeroException)
            {
                DGVCeldas[colIndex, rowIndex].Value = "#¡DIV/O!";
                MessageBox.Show("Error: División entre cero.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception)
            {
                DGVCeldas[colIndex, rowIndex].Value = "#VALOR";
            }
        }

        private string EvaluarFormula(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                throw new ArgumentException("Fórmula vacía");

            formula = formula.Trim();

            if (formula.StartsWith("="))
                formula = formula.Substring(1);

            formula = EvaluarFunciones(formula);
            formula = ReemplazarReferencias(formula);

            if (formula.StartsWith("\"") && formula.EndsWith("\""))
            {
                return formula.Substring(1, formula.Length - 2);
            }

            DataTable dt = new DataTable();
            var resultado = dt.Compute(formula, "");
            return Convert.ToDouble(resultado).ToString(CultureInfo.InvariantCulture);
        }

        private string EvaluarFunciones(string formula)
        {
            string[] funciones = { "SUMA", "RESTA", "MULTI", "DIVIDIR", "MIN", "MAX", "PROM", "CONTVAL", "CONCAT", "POWER", "SQRT", "MODA", "ABS"};

            foreach (string f in funciones)
            {
                string pattern = $@"(?i){f}\s*\(([^()]*)\)";

                while (true)
                {
                    var match = Regex.Match(formula, pattern);
                    if (!match.Success) break;

                    string argsStr = match.Groups[1].Value;

                    if (f.ToUpper() == "CONCAT")
                    {
                        List<string> textos = new List<string>();
                        string[] argsConcat = argsStr.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                        foreach (string arg in argsConcat)
                        {
                            string a = arg.Trim().ToUpper();

                            if (a.Contains(":"))
                            {
                                string[] partes = a.Split(':');
                                if (partes.Length == 2)
                                    textos.AddRange(ObtenerValoresRangoComoTexto(partes[0], partes[1]));
                            }
                            else if (char.IsLetter(a[0]))
                            {
                                var val = ObtenerTextoReferencia(a);
                                if (!string.IsNullOrWhiteSpace(val))
                                    textos.Add(val);
                            }
                            else
                            {
                                textos.Add(arg.Trim());
                            }
                        }

                        string resultadoTexto = string.Join(" ", textos);

                        formula = formula.Substring(0, match.Index)
                                + "\"" + resultadoTexto + "\""
                                + formula.Substring(match.Index + match.Length);

                        continue;
                    }

                    double resultado = 0;
                    List<double> valores = new List<double>();

                    string[] args = argsStr.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string arg in args)
                    {
                        string a = arg.Trim().ToUpper();

                        if (a.Contains(":"))
                        {
                            string[] partes = a.Split(':');
                            if (partes.Length == 2)
                                valores.AddRange(ObtenerValoresRango(partes[0], partes[1]));
                        }
                        else if (char.IsLetter(a[0]))
                        {
                            valores.Add(ObtenerValorReferencia(a));
                        }
                        else
                        {
                            if (!double.TryParse(a, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                                val = 0;
                            valores.Add(val);
                        }
                    }

                    switch (f.ToUpper())
                    {
                        case "SUMA":
                            resultado = valores.Sum();
                            break;

                        case "RESTA":
                            if (valores.Count > 0)
                            {
                                resultado = valores[0];
                                for (int i = 1; i < valores.Count; i++)
                                    resultado -= valores[i];
                            }
                            break;

                        case "MULTI":
                            resultado = 1;
                            foreach (var v in valores)
                                resultado *= v;
                            break;

                        case "DIVIDIR":
                            if (valores.Count > 0)
                            {
                                resultado = valores[0];
                                for (int i = 1; i < valores.Count; i++)
                                {
                                    if (valores[i] == 0)
                                        throw new DivideByZeroException();
                                    resultado /= valores[i];
                                }
                            }
                            break;

                        case "MIN":
                            if (valores.Count > 0)
                                resultado = valores.Min();
                            break;

                        case "MAX":
                            if (valores.Count > 0)
                                resultado = valores.Max();
                            break;

                        case "PROM":
                            if (valores.Count > 0)
                                resultado = valores.Average();
                            break;

                        case "CONTVAL":
                            resultado = valores.Count(v => !double.IsNaN(v));
                            break;

                        case "POWER":
                            if (valores.Count >= 2)
                                resultado = Math.Pow(valores[0], valores[1]);
                            break;

                        case "SQRT":
                            if (valores.Count >= 1)
                                resultado = Math.Sqrt(valores[0]);
                            break;

                        case "MODA":
                            if (valores.Count > 0)
                                resultado = valores.GroupBy(v => v)
                                                    .OrderByDescending(g => g.Count())
                                                    .ThenBy(g => g.Key)
                                                    .First().Key;
                            break;

                        case "ABS":
                            if (valores.Count >= 1)
                                resultado = Math.Abs(valores[0]);
                            break;

                    }

                    formula = formula.Substring(0, match.Index)
                            + resultado.ToString(CultureInfo.InvariantCulture)
                            + formula.Substring(match.Index + match.Length);
                }
            }

            return formula;
        }

        private string ReemplazarReferencias(string formula)
        {
            string pattern = @"([A-Z]+[0-9]+)";
            return Regex.Replace(formula, pattern, match =>
            {
                string referencia = match.Value;
                double val = ObtenerValorReferencia(referencia);
                return val.ToString(CultureInfo.InvariantCulture);
            });
        }

        private double ObtenerValorReferencia(string referencia)
        {
            int i = 0;
            while (i < referencia.Length && char.IsLetter(referencia[i])) i++;

            string colPart = referencia.Substring(0, i).ToUpper();
            string rowPart = referencia.Substring(i);

            int colIndex = ReferenciasCeldas.ObtenerNumColumna(colPart);

            if (!int.TryParse(rowPart, out int rowIndex))
                return 0;

            rowIndex -= 1;

            if (colIndex < 0 || colIndex >= DGVCeldas.Columns.Count ||
                rowIndex < 0 || rowIndex >= DGVCeldas.Rows.Count)
                return 0;

            var celdaValor = DGVCeldas[colIndex, rowIndex].Value;

            if (celdaValor == null)
                return 0;

            if (!double.TryParse(celdaValor.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double valorNumerico))
                return 0;

            return valorNumerico;
        }

        private List<double> ObtenerValoresRango(string inicio, string fin)
        {
            int i = 0;

            while (i < inicio.Length && char.IsLetter(inicio[i])) i++;
            string colIni = inicio.Substring(0, i);
            string rowIni = inicio.Substring(i);

            i = 0;
            while (i < fin.Length && char.IsLetter(fin[i])) i++;
            string colFin = fin.Substring(0, i);
            string rowFin = fin.Substring(i);

            int colStart = ReferenciasCeldas.ObtenerNumColumna(colIni);
            int colEnd = ReferenciasCeldas.ObtenerNumColumna(colFin);

            if (!int.TryParse(rowIni, out int rowStart)) rowStart = 1;
            if (!int.TryParse(rowFin, out int rowEnd)) rowEnd = 1;

            rowStart -= 1; rowEnd -= 1;

            if (colStart > colEnd) (colStart, colEnd) = (colEnd, colStart);
            if (rowStart > rowEnd) (rowStart, rowEnd) = (rowEnd, rowStart);

            var valores = new List<double>();

            for (int c = colStart; c <= colEnd; c++)
            {
                for (int r = rowStart; r <= rowEnd; r++)
                {
                    if (c < DGVCeldas.Columns.Count && r < DGVCeldas.Rows.Count && c >= 0 && r >= 0)
                        valores.Add(ObtenerValorReferencia(ReferenciasCeldas.ObtenerCeldaActual(c) + (r + 1)));
                }
            }

            return valores;
        }

        private List<string> ObtenerValoresRangoComoTexto(string inicio, string fin)
        {
            List<string> textos = new List<string>();

            int i = 0;
            while (i < inicio.Length && char.IsLetter(inicio[i])) i++;
            string colIni = inicio.Substring(0, i);
            string rowIni = inicio.Substring(i);

            i = 0;
            while (i < fin.Length && char.IsLetter(fin[i])) i++;
            string colFin = fin.Substring(0, i);
            string rowFin = fin.Substring(i);

            int colStart = ReferenciasCeldas.ObtenerNumColumna(colIni);
            int colEnd = ReferenciasCeldas.ObtenerNumColumna(colFin);

            if (!int.TryParse(rowIni, out int rowStart)) rowStart = 1;
            if (!int.TryParse(rowFin, out int rowEnd)) rowEnd = 1;

            rowStart -= 1; rowEnd -= 1;

            if (colStart > colEnd) (colStart, colEnd) = (colEnd, colStart);
            if (rowStart > rowEnd) (rowStart, rowEnd) = (rowEnd, rowStart);

            for (int c = colStart; c <= colEnd; c++)
            {
                for (int r = rowStart; r <= rowEnd; r++)
                {
                    if (c < DGVCeldas.Columns.Count && r < DGVCeldas.Rows.Count && c >= 0 && r >= 0)
                    {
                        var val = DGVCeldas[c, r].Value;
                        if (val != null && !string.IsNullOrWhiteSpace(val.ToString()))
                            textos.Add(val.ToString());
                    }
                }
            }

            return textos;
        }

        private string ObtenerTextoReferencia(string referencia)
        {
            int i = 0;
            while (i < referencia.Length && char.IsLetter(referencia[i])) i++;

            string colPart = referencia.Substring(0, i).ToUpper();
            string rowPart = referencia.Substring(i);

            int colIndex = ReferenciasCeldas.ObtenerNumColumna(colPart);

            if (!int.TryParse(rowPart, out int rowIndex))
                return string.Empty;

            rowIndex -= 1;

            if (colIndex < 0 || colIndex >= DGVCeldas.Columns.Count ||
                rowIndex < 0 || rowIndex >= DGVCeldas.Rows.Count)
                return string.Empty;

            var val = DGVCeldas[colIndex, rowIndex].Value;
            return val?.ToString() ?? string.Empty;
        }

        public void EvaluarOperacionSeleccionada(DataGridView dgv, string operacion)
        {
            var celdasSeleccionadas = dgv.SelectedCells.Cast<DataGridViewCell>()
                .OrderBy(c => c.RowIndex)
                .ThenBy(c => c.ColumnIndex)
                .ToList();

            if (!celdasSeleccionadas.Any())
            {
                MessageBox.Show("No hay celdas seleccionadas.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var celdaResultado = celdasSeleccionadas.FirstOrDefault(c => c.Value == null || string.IsNullOrWhiteSpace(c.Value.ToString()));
            if (celdaResultado == null)
            {
                MessageBox.Show("Debe seleccionar una celda vacía para mostrar el resultado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var celdasArgumentos = celdasSeleccionadas
                .Where(c => c != celdaResultado)
                .Where(c => c.Value != null && !string.IsNullOrWhiteSpace(c.Value.ToString()))
                .ToList();

            if (!celdasArgumentos.Any())
            {
                MessageBox.Show("No hay valores para operar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if ((operacion == "SQRT" || operacion == "ABS" || operacion == "POWER") && celdasArgumentos.Count != 1)
            {
                MessageBox.Show($"{operacion} solo necesita que seleccione una celda con valor y una vacía para el resultado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string formula;

            if (operacion == "SQRT" || operacion == "ABS")
            {
                string referencia = ReferenciasCeldas.ObtenerCeldaActual(celdasArgumentos[0].ColumnIndex) + (celdasArgumentos[0].RowIndex + 1);
                formula = $"={operacion}({referencia})";
            }
            else if (operacion == "POWER")
            {
                string referencia = ReferenciasCeldas.ObtenerCeldaActual(celdasArgumentos[0].ColumnIndex) + (celdasArgumentos[0].RowIndex + 1);
                formula = $"=POWER({referencia},2)";
            }
            else
            {
                var referencias = celdasArgumentos.Select(c =>
                {
                    string col = ReferenciasCeldas.ObtenerCeldaActual(c.ColumnIndex);
                    string row = (c.RowIndex + 1).ToString();
                    return $"{col}{row}";
                });

                formula = $"={operacion}({string.Join(",", referencias)})";
            }

            try
            {
                EvaluarYAsignar(celdaResultado.ColumnIndex, celdaResultado.RowIndex, formula);
            }
            catch (DivideByZeroException)
            {
                dgv[celdaResultado.ColumnIndex, celdaResultado.RowIndex].Value = "#¡DIV/O!";
                MessageBox.Show("Error: División entre cero no permitida.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception)
            {
                dgv[celdaResultado.ColumnIndex, celdaResultado.RowIndex].Value = "#VALOR";
                MessageBox.Show("Error al evaluar la fórmula.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}