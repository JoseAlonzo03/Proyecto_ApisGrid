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
    public partial class LogoIniciando : Form
    {

        int tiempoTotal = 3000;
        int intervalo = 100;
        int progreso = 0;

        public LogoIniciando()
        {
            InitializeComponent();
            progressBar1.Minimum = 0;
            progressBar1.Maximum = tiempoTotal;
            progressBar1.Value = 0;

            timer1.Interval = intervalo;
            timer1.Start();
        }

        private void LogoIniciando_Load(object sender, EventArgs e)
        {
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progreso += intervalo;
            if (progreso <= tiempoTotal)
            {
                progressBar1.Value = progreso;
            }
            else
            {
                timer1.Stop();
                Libro1 libro1 = new Libro1();
                libro1.Show();
                this.Hide();
            }
        }
    }
}
