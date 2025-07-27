using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProyectodeTeoríadelaComputación
{
    internal class ReferenciasCeldas
    {
        public static string ObtenerCeldaActual(int colIndex)
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

        public static int ObtenerNumColumna(string colLetters)
        {
            int col = 0;
            for (int i = 0; i < colLetters.Length; i++)
            {
                col *= 26;
                col += (colLetters[i] - 'A' + 1);
            }
            return col - 1;
        }
    }
}
