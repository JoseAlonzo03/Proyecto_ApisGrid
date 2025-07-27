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
    }
}
