using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Drawing;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Windows.Forms;

namespace ProyectodeTeoríadelaComputación
{
    internal class GuardarDatos
    {
        public static void ExportarExcel(DataGridView dgv)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Archivo Excel|*.xlsx", FileName = "TablaExportada.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Datos");

                        for (int r = 0; r < dgv.Rows.Count; r++)
                        {
                            var row = dgv.Rows[r];
                            if (row.IsNewRow) continue;

                            for (int c = 0; c < dgv.Columns.Count; c++)
                            {
                                var dgvCell = row.Cells[c];
                                var cell = worksheet.Cell(r + 1, c + 1);

                                if (dgvCell.Value != null)
                                {
                                    string valorStr = dgvCell.Value.ToString();

                                    if (double.TryParse(valorStr, out double num))
                                    {
                                        cell.Value = num;
                                    }
                                    else if (DateTime.TryParse(valorStr, out DateTime fecha))
                                    {
                                        cell.Value = fecha;
                                    }
                                    else
                                    {
                                        cell.Value = valorStr;
                                    }
                                }
                                else
                                {
                                    cell.Clear();
                                }

                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                                Color backColor = dgvCell.Style.BackColor;
                                if (backColor.IsEmpty)
                                {
                                    backColor = dgvCell.OwningColumn.DefaultCellStyle.BackColor;
                                    if (backColor.IsEmpty)
                                        backColor = dgv.DefaultCellStyle.BackColor;
                                }
                                cell.Style.Fill.BackgroundColor = XLColor.FromColor(backColor);

                                Color foreColor = dgvCell.Style.ForeColor;
                                if (foreColor.IsEmpty)
                                    foreColor = dgv.DefaultCellStyle.ForeColor;
                                cell.Style.Font.FontColor = XLColor.FromColor(foreColor);

                                var font = dgvCell.Style.Font ?? dgv.DefaultCellStyle.Font;
                                if (font != null)
                                {
                                    cell.Style.Font.FontName = font.FontFamily.Name;
                                    cell.Style.Font.FontSize = font.Size;
                                    cell.Style.Font.Bold = font.Bold;
                                    cell.Style.Font.Italic = font.Italic;
                                    cell.Style.Font.Underline = font.Underline ? XLFontUnderlineValues.Single : XLFontUnderlineValues.None;
                                }

                                var align = dgvCell.Style.Alignment;
                                if (align == DataGridViewContentAlignment.NotSet)
                                    align = dgvCell.OwningColumn.DefaultCellStyle.Alignment;

                                switch (align)
                                {
                                    case DataGridViewContentAlignment.MiddleCenter:
                                    case DataGridViewContentAlignment.TopCenter:
                                    case DataGridViewContentAlignment.BottomCenter:
                                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                        break;
                                    case DataGridViewContentAlignment.MiddleRight:
                                    case DataGridViewContentAlignment.TopRight:
                                    case DataGridViewContentAlignment.BottomRight:
                                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                                        break;
                                    default:
                                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                                        break;
                                }
                                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            }
                        }

                        for (int c = 0; c < dgv.Columns.Count; c++)
                        {
                            int anchoPixel = dgv.Columns[c].Width;

                            double anchoExcel = anchoPixel / 7.0;

                            worksheet.Column(c + 1).Width = anchoExcel;
                        }

                        workbook.SaveAs(sfd.FileName);
                        MessageBox.Show("La tabla se guardó como archivo Excel con éxito.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        public static void ExportarPDF(DataGridView dgv)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Documento PDF|*.pdf", FileName = "TablaExportada.pdf" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        Document doc = new Document(PageSize.A4.Rotate(), 10, 10, 50, 50);
                        PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));

                        writer.PageEvent = new PdfPageEvents();

                        doc.Open();

                        System.Drawing.Image img = ProyectodeTeoríadelaComputación.Properties.Resources.ApisGrid_Logo;
                        using (var ms = new MemoryStream())
                        {
                            img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                            iTextSharp.text.Image pdfImg = iTextSharp.text.Image.GetInstance(ms.ToArray());
                            pdfImg.Alignment = Element.ALIGN_CENTER;
                            pdfImg.ScaleToFit(200f, 100f);
                            doc.Add(pdfImg);
                        }

                        doc.Add(new Paragraph("\n"));

                        PdfPTable table = new PdfPTable(dgv.Columns.Count + 1);
                        table.WidthPercentage = 100;

                        PdfPCell headerCell = new PdfPCell(new Phrase("No.",
                            ConvertirFuente(dgv.RowHeadersDefaultCellStyle.Font ?? dgv.Font,
                                            dgv.RowHeadersDefaultCellStyle.ForeColor.IsEmpty ? System.Drawing.Color.Black : dgv.RowHeadersDefaultCellStyle.ForeColor)));
                        headerCell.BackgroundColor = BaseColor.LIGHT_GRAY;
                        headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
                        table.AddCell(headerCell);

                        foreach (DataGridViewColumn col in dgv.Columns)
                        {
                            PdfPCell cell = new PdfPCell(new Phrase(col.HeaderText,
                                ConvertirFuente(col.HeaderCell.Style.Font ?? dgv.ColumnHeadersDefaultCellStyle.Font ?? dgv.Font,
                                                col.HeaderCell.Style.ForeColor.IsEmpty ? System.Drawing.Color.Black : col.HeaderCell.Style.ForeColor)));
                            cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            table.AddCell(cell);
                        }

                        int filaNumero = 1;
                        foreach (DataGridViewRow row in dgv.Rows)
                        {
                            if (row.IsNewRow) continue;

                            PdfPCell cellNum = new PdfPCell(new Phrase(filaNumero.ToString(),
                                ConvertirFuente(dgv.RowHeadersDefaultCellStyle.Font ?? dgv.Font,
                                                dgv.RowHeadersDefaultCellStyle.ForeColor.IsEmpty ? System.Drawing.Color.Black : dgv.RowHeadersDefaultCellStyle.ForeColor)));
                            cellNum.HorizontalAlignment = Element.ALIGN_CENTER;
                            cellNum.BackgroundColor = BaseColor.LIGHT_GRAY;
                            table.AddCell(cellNum);
                            filaNumero++;

                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                System.Drawing.Color backColor = cell.Style.BackColor;
                                if (backColor.IsEmpty)
                                {
                                    backColor = cell.OwningColumn.DefaultCellStyle.BackColor;
                                    if (backColor.IsEmpty)
                                        backColor = dgv.DefaultCellStyle.BackColor;
                                }
                                BaseColor baseColor = new BaseColor(backColor.R, backColor.G, backColor.B);

                                System.Drawing.Font cellFont = cell.Style.Font ?? dgv.DefaultCellStyle.Font ?? dgv.Font;
                                System.Drawing.Color foreColor = cell.Style.ForeColor;
                                if (foreColor.IsEmpty)
                                    foreColor = dgv.DefaultCellStyle.ForeColor;

                                iTextSharp.text.Font pdfFont = ConvertirFuente(cellFont, foreColor);

                                PdfPCell dataCell = new PdfPCell(new Phrase(cell.Value?.ToString() ?? "", pdfFont));

                                DataGridViewContentAlignment align = cell.Style.Alignment;
                                if (align == DataGridViewContentAlignment.NotSet)
                                    align = cell.OwningColumn.DefaultCellStyle.Alignment;

                                dataCell.HorizontalAlignment = MapearAlineacion(align);

                                dataCell.BackgroundColor = baseColor;

                                table.AddCell(dataCell);
                            }
                        }

                        doc.Add(table);
                        doc.Close();

                        MessageBox.Show("La tabla se guardó como archivo PDF con éxito.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al guardar como PDF: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private static iTextSharp.text.Font ConvertirFuente(System.Drawing.Font fuente, System.Drawing.Color colorTexto)
        {
            int estilo = iTextSharp.text.Font.NORMAL;

            if (fuente.Bold) estilo |= iTextSharp.text.Font.BOLD;
            if (fuente.Italic) estilo |= iTextSharp.text.Font.ITALIC;
            if (fuente.Underline) estilo |= iTextSharp.text.Font.UNDERLINE;

            BaseColor baseColorTexto = new BaseColor(colorTexto.R, colorTexto.G, colorTexto.B);

            string fontName = BaseFont.HELVETICA;

            string nombreFuente = fuente.FontFamily.Name.ToLower();
            if (nombreFuente.Contains("courier"))
                fontName = BaseFont.COURIER;
            else if (nombreFuente.Contains("times"))
                fontName = BaseFont.TIMES_ROMAN;
            else if (nombreFuente.Contains("helvetica") || nombreFuente.Contains("arial"))
                fontName = BaseFont.HELVETICA;

            BaseFont baseFont = BaseFont.CreateFont(fontName, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            return new iTextSharp.text.Font(baseFont, fuente.Size, estilo, baseColorTexto);
        }

        private static int MapearAlineacion(DataGridViewContentAlignment align)
        {
            switch (align)
            {
                case DataGridViewContentAlignment.BottomCenter:
                case DataGridViewContentAlignment.MiddleCenter:
                case DataGridViewContentAlignment.TopCenter:
                    return Element.ALIGN_CENTER;
                case DataGridViewContentAlignment.BottomRight:
                case DataGridViewContentAlignment.MiddleRight:
                case DataGridViewContentAlignment.TopRight:
                    return Element.ALIGN_RIGHT;
                default:
                    return Element.ALIGN_LEFT;
            }
        }

        public class PdfPageEvents : PdfPageEventHelper
        {
            public override void OnEndPage(PdfWriter writer, Document document)
            {
                var fechaHora = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                var cb = writer.DirectContent;
                BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                cb.BeginText();
                cb.SetFontAndSize(bf, 10);
                float x = document.Left + 40;
                float y = document.Bottom - 20;
                cb.SetTextMatrix(x, y);
                cb.ShowText(fechaHora);
                cb.EndText();
            }
        }

        public static void ImportarExcel(DataGridView dgv, string rutaArchivo)
        {
            try
            {
                using (var workbook = new XLWorkbook(rutaArchivo))
                {
                    var worksheet = workbook.Worksheets.First();

                    foreach (DataGridViewRow row in dgv.Rows)
                    {
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            cell.Value = null;
                            cell.Style.BackColor = Color.Empty;
                            cell.Style.ForeColor = Color.Empty;
                        }
                    }

                    int colCount = worksheet.LastColumnUsed().ColumnNumber();
                    int rowCount = worksheet.LastRowUsed().RowNumber();

                    for (int c = dgv.Columns.Count; c < colCount; c++)
                    {
                        string colName = ReferenciasCeldas.ObtenerCeldaActual(c);
                        dgv.Columns.Add(colName, colName);
                        dgv.Columns[c].Width = 80;
                    }

                    int filasFaltantes = rowCount - dgv.Rows.Count;
                    if (filasFaltantes > 0)
                        dgv.Rows.Add(filasFaltantes);

                    for (int r = 1; r <= rowCount; r++)
                    {
                        for (int c = 1; c <= colCount; c++)
                        {
                            var xlCell = worksheet.Cell(r, c);

                            dgv[c - 1, r - 1].Value = xlCell.Value.ToString();

                            var bgColor = xlCell.Style.Fill.BackgroundColor;
                            if (bgColor.ColorType == XLColorType.Color)
                            {
                                var sysColor = Color.FromArgb(bgColor.Color.ToArgb());
                                dgv[c - 1, r - 1].Style.BackColor = sysColor;
                            }

                            var fgColor = xlCell.Style.Font.FontColor;
                            if (fgColor.ColorType == XLColorType.Color)
                            {
                                var sysForeColor = Color.FromArgb(fgColor.Color.ToArgb());
                                dgv[c - 1, r - 1].Style.ForeColor = sysForeColor;
                            }
                        }
                    }
                }

                MessageBox.Show("Archivo Excel cargado correctamente.", "Éxito",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir el archivo: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}