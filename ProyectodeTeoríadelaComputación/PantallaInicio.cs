using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProyectodeTeoríadelaComputación
{
    public partial class PantallaInicio : Form
    {
        public PantallaInicio()
        {
            InitializeComponent();
        }

        private void PantallaInicio_Load(object sender, EventArgs e)
        {
          
        }

        private void btnNuevo_Click(object sender, EventArgs e)
        {
            Libro1 libro1 = new Libro1();
            libro1.Show();
            this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                Filter = "Archivos Excel (*.xlsx)|*.xlsx",
                Title = "Abrir archivo Excel"
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string ruta = ofd.FileName;

                Libro1 libro1;

                if (Application.OpenForms["Libro1"] != null)
                {
                    libro1 = (Libro1)Application.OpenForms["Libro1"];
                }
                else
                {
                    libro1 = new Libro1();
                }

                libro1.Show();
                this.Hide();

                libro1.CargarArchivoExcel(ruta);
            }
        }
    }
}