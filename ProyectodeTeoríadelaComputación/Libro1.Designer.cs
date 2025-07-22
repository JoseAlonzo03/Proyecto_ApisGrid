namespace ProyectodeTeoríadelaComputación
{
    partial class Libro1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Libro1));
            this.DGVCeldas = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnNegrita = new System.Windows.Forms.Button();
            this.btnCursiva = new System.Windows.Forms.Button();
            this.cmbFuentes = new System.Windows.Forms.ComboBox();
            this.cmbTamaño = new System.Windows.Forms.ComboBox();
            this.btnSubrayado = new System.Windows.Forms.Button();
            this.txtCeldaActiva = new System.Windows.Forms.TextBox();
            this.GBAlineación = new System.Windows.Forms.GroupBox();
            this.btnTXAbajo = new System.Windows.Forms.Button();
            this.btnTXArriba = new System.Windows.Forms.Button();
            this.btnTXCentroV = new System.Windows.Forms.Button();
            this.btnTXDerecha = new System.Windows.Forms.Button();
            this.btnTXIzquierda = new System.Windows.Forms.Button();
            this.btnTXCentro = new System.Windows.Forms.Button();
            this.GBFuentes = new System.Windows.Forms.GroupBox();
            this.btnRellenoColor = new System.Windows.Forms.Button();
            this.lblFuentes = new System.Windows.Forms.Label();
            this.lblAlineación = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.DGVCeldas)).BeginInit();
            this.GBAlineación.SuspendLayout();
            this.GBFuentes.SuspendLayout();
            this.SuspendLayout();
            // 
            // DGVCeldas
            // 
            this.DGVCeldas.AllowUserToAddRows = false;
            this.DGVCeldas.AllowUserToOrderColumns = true;
            this.DGVCeldas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DGVCeldas.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DGVCeldas.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.DGVCeldas.BackgroundColor = System.Drawing.Color.White;
            this.DGVCeldas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGVCeldas.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3,
            this.Column4,
            this.Column5,
            this.Column6,
            this.Column7,
            this.Column8,
            this.Column9,
            this.Column10});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.DGVCeldas.DefaultCellStyle = dataGridViewCellStyle1;
            this.DGVCeldas.EnableHeadersVisualStyles = false;
            this.DGVCeldas.Location = new System.Drawing.Point(12, 203);
            this.DGVCeldas.Name = "DGVCeldas";
            this.DGVCeldas.RowHeadersWidth = 51;
            this.DGVCeldas.RowTemplate.Height = 24;
            this.DGVCeldas.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.DGVCeldas.Size = new System.Drawing.Size(1284, 331);
            this.DGVCeldas.TabIndex = 0;
            this.DGVCeldas.CurrentCellChanged += new System.EventHandler(this.DGVCeldas_CurrentCellChanged);
            this.DGVCeldas.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.DGVCeldas_RowPostPaint);
            this.DGVCeldas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.DGVCeldas_KeyDown);
            // 
            // Column1
            // 
            this.Column1.HeaderText = "A";
            this.Column1.MinimumWidth = 6;
            this.Column1.Name = "Column1";
            // 
            // Column2
            // 
            this.Column2.HeaderText = "B";
            this.Column2.MinimumWidth = 6;
            this.Column2.Name = "Column2";
            // 
            // Column3
            // 
            this.Column3.HeaderText = "C";
            this.Column3.MinimumWidth = 6;
            this.Column3.Name = "Column3";
            // 
            // Column4
            // 
            this.Column4.HeaderText = "D";
            this.Column4.MinimumWidth = 6;
            this.Column4.Name = "Column4";
            // 
            // Column5
            // 
            this.Column5.HeaderText = "E";
            this.Column5.MinimumWidth = 6;
            this.Column5.Name = "Column5";
            // 
            // Column6
            // 
            this.Column6.HeaderText = "F";
            this.Column6.MinimumWidth = 6;
            this.Column6.Name = "Column6";
            // 
            // Column7
            // 
            this.Column7.HeaderText = "G";
            this.Column7.MinimumWidth = 6;
            this.Column7.Name = "Column7";
            // 
            // Column8
            // 
            this.Column8.HeaderText = "H";
            this.Column8.MinimumWidth = 6;
            this.Column8.Name = "Column8";
            // 
            // Column9
            // 
            this.Column9.HeaderText = "I";
            this.Column9.MinimumWidth = 6;
            this.Column9.Name = "Column9";
            // 
            // Column10
            // 
            this.Column10.HeaderText = "J";
            this.Column10.MinimumWidth = 6;
            this.Column10.Name = "Column10";
            // 
            // btnNegrita
            // 
            this.btnNegrita.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNegrita.Location = new System.Drawing.Point(6, 104);
            this.btnNegrita.Name = "btnNegrita";
            this.btnNegrita.Size = new System.Drawing.Size(45, 45);
            this.btnNegrita.TabIndex = 1;
            this.btnNegrita.Text = "N";
            this.btnNegrita.UseVisualStyleBackColor = true;
            this.btnNegrita.Click += new System.EventHandler(this.btnNegrita_Click);
            // 
            // btnCursiva
            // 
            this.btnCursiva.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCursiva.Location = new System.Drawing.Point(57, 104);
            this.btnCursiva.Name = "btnCursiva";
            this.btnCursiva.Size = new System.Drawing.Size(45, 45);
            this.btnCursiva.TabIndex = 2;
            this.btnCursiva.Text = "K";
            this.btnCursiva.UseVisualStyleBackColor = true;
            this.btnCursiva.Click += new System.EventHandler(this.btnCursiva_Click);
            // 
            // cmbFuentes
            // 
            this.cmbFuentes.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbFuentes.FormattingEnabled = true;
            this.cmbFuentes.Location = new System.Drawing.Point(6, 65);
            this.cmbFuentes.Name = "cmbFuentes";
            this.cmbFuentes.Size = new System.Drawing.Size(230, 33);
            this.cmbFuentes.TabIndex = 3;
            this.cmbFuentes.SelectedIndexChanged += new System.EventHandler(this.cmbFuentes_SelectedIndexChanged);
            // 
            // cmbTamaño
            // 
            this.cmbTamaño.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbTamaño.FormattingEnabled = true;
            this.cmbTamaño.Location = new System.Drawing.Point(159, 108);
            this.cmbTamaño.Name = "cmbTamaño";
            this.cmbTamaño.Size = new System.Drawing.Size(78, 37);
            this.cmbTamaño.TabIndex = 4;
            this.cmbTamaño.SelectedIndexChanged += new System.EventHandler(this.cmbTamaño_SelectedIndexChanged);
            // 
            // btnSubrayado
            // 
            this.btnSubrayado.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSubrayado.Location = new System.Drawing.Point(108, 104);
            this.btnSubrayado.Name = "btnSubrayado";
            this.btnSubrayado.Size = new System.Drawing.Size(45, 45);
            this.btnSubrayado.TabIndex = 5;
            this.btnSubrayado.Text = "S";
            this.btnSubrayado.UseVisualStyleBackColor = true;
            this.btnSubrayado.Click += new System.EventHandler(this.btnSubrayado_Click);
            // 
            // txtCeldaActiva
            // 
            this.txtCeldaActiva.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCeldaActiva.Location = new System.Drawing.Point(6, 21);
            this.txtCeldaActiva.Name = "txtCeldaActiva";
            this.txtCeldaActiva.Size = new System.Drawing.Size(86, 34);
            this.txtCeldaActiva.TabIndex = 6;
            this.txtCeldaActiva.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCeldaActiva_KeyDown);
            // 
            // GBAlineación
            // 
            this.GBAlineación.Controls.Add(this.btnTXAbajo);
            this.GBAlineación.Controls.Add(this.btnTXArriba);
            this.GBAlineación.Controls.Add(this.btnTXCentroV);
            this.GBAlineación.Controls.Add(this.btnTXDerecha);
            this.GBAlineación.Controls.Add(this.btnTXIzquierda);
            this.GBAlineación.Controls.Add(this.btnTXCentro);
            this.GBAlineación.Location = new System.Drawing.Point(278, 12);
            this.GBAlineación.Name = "GBAlineación";
            this.GBAlineación.Size = new System.Drawing.Size(160, 160);
            this.GBAlineación.TabIndex = 11;
            this.GBAlineación.TabStop = false;
            // 
            // btnTXAbajo
            // 
            this.btnTXAbajo.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnTXAbajo.BackgroundImage")));
            this.btnTXAbajo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnTXAbajo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTXAbajo.Location = new System.Drawing.Point(108, 49);
            this.btnTXAbajo.Name = "btnTXAbajo";
            this.btnTXAbajo.Size = new System.Drawing.Size(45, 45);
            this.btnTXAbajo.TabIndex = 12;
            this.btnTXAbajo.UseVisualStyleBackColor = true;
            this.btnTXAbajo.Click += new System.EventHandler(this.btnTXAbajo_Click);
            // 
            // btnTXArriba
            // 
            this.btnTXArriba.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnTXArriba.BackgroundImage")));
            this.btnTXArriba.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnTXArriba.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTXArriba.Location = new System.Drawing.Point(6, 49);
            this.btnTXArriba.Name = "btnTXArriba";
            this.btnTXArriba.Size = new System.Drawing.Size(45, 45);
            this.btnTXArriba.TabIndex = 10;
            this.btnTXArriba.UseVisualStyleBackColor = true;
            this.btnTXArriba.Click += new System.EventHandler(this.btnTXArriba_Click);
            // 
            // btnTXCentroV
            // 
            this.btnTXCentroV.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnTXCentroV.BackgroundImage")));
            this.btnTXCentroV.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnTXCentroV.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTXCentroV.Location = new System.Drawing.Point(57, 49);
            this.btnTXCentroV.Name = "btnTXCentroV";
            this.btnTXCentroV.Size = new System.Drawing.Size(45, 45);
            this.btnTXCentroV.TabIndex = 11;
            this.btnTXCentroV.UseVisualStyleBackColor = true;
            this.btnTXCentroV.Click += new System.EventHandler(this.btnTXCentroV_Click);
            // 
            // btnTXDerecha
            // 
            this.btnTXDerecha.BackgroundImage = global::ProyectodeTeoríadelaComputación.Properties.Resources.txtDerecha;
            this.btnTXDerecha.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnTXDerecha.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTXDerecha.Location = new System.Drawing.Point(108, 100);
            this.btnTXDerecha.Name = "btnTXDerecha";
            this.btnTXDerecha.Size = new System.Drawing.Size(45, 45);
            this.btnTXDerecha.TabIndex = 9;
            this.btnTXDerecha.UseVisualStyleBackColor = true;
            this.btnTXDerecha.Click += new System.EventHandler(this.btnTXDerecha_Click);
            // 
            // btnTXIzquierda
            // 
            this.btnTXIzquierda.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnTXIzquierda.BackgroundImage")));
            this.btnTXIzquierda.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnTXIzquierda.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTXIzquierda.Location = new System.Drawing.Point(6, 100);
            this.btnTXIzquierda.Name = "btnTXIzquierda";
            this.btnTXIzquierda.Size = new System.Drawing.Size(45, 45);
            this.btnTXIzquierda.TabIndex = 7;
            this.btnTXIzquierda.UseVisualStyleBackColor = true;
            this.btnTXIzquierda.Click += new System.EventHandler(this.btnTXIzquierda_Click);
            // 
            // btnTXCentro
            // 
            this.btnTXCentro.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnTXCentro.BackgroundImage")));
            this.btnTXCentro.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnTXCentro.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTXCentro.Location = new System.Drawing.Point(57, 100);
            this.btnTXCentro.Name = "btnTXCentro";
            this.btnTXCentro.Size = new System.Drawing.Size(45, 45);
            this.btnTXCentro.TabIndex = 8;
            this.btnTXCentro.UseVisualStyleBackColor = true;
            this.btnTXCentro.Click += new System.EventHandler(this.btnTXCentro_Click);
            // 
            // GBFuentes
            // 
            this.GBFuentes.Controls.Add(this.btnRellenoColor);
            this.GBFuentes.Controls.Add(this.cmbTamaño);
            this.GBFuentes.Controls.Add(this.btnNegrita);
            this.GBFuentes.Controls.Add(this.txtCeldaActiva);
            this.GBFuentes.Controls.Add(this.btnCursiva);
            this.GBFuentes.Controls.Add(this.btnSubrayado);
            this.GBFuentes.Controls.Add(this.cmbFuentes);
            this.GBFuentes.Location = new System.Drawing.Point(12, 12);
            this.GBFuentes.Name = "GBFuentes";
            this.GBFuentes.Size = new System.Drawing.Size(242, 160);
            this.GBFuentes.TabIndex = 12;
            this.GBFuentes.TabStop = false;
            // 
            // btnRellenoColor
            // 
            this.btnRellenoColor.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnRellenoColor.BackgroundImage")));
            this.btnRellenoColor.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnRellenoColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRellenoColor.Location = new System.Drawing.Point(98, 15);
            this.btnRellenoColor.Name = "btnRellenoColor";
            this.btnRellenoColor.Size = new System.Drawing.Size(45, 45);
            this.btnRellenoColor.TabIndex = 7;
            this.btnRellenoColor.UseVisualStyleBackColor = true;
            this.btnRellenoColor.Click += new System.EventHandler(this.btnRellenoColor_Click);
            // 
            // lblFuentes
            // 
            this.lblFuentes.AutoSize = true;
            this.lblFuentes.Location = new System.Drawing.Point(100, 175);
            this.lblFuentes.Name = "lblFuentes";
            this.lblFuentes.Size = new System.Drawing.Size(55, 16);
            this.lblFuentes.TabIndex = 13;
            this.lblFuentes.Text = "Fuentes";
            // 
            // lblAlineación
            // 
            this.lblAlineación.AutoSize = true;
            this.lblAlineación.Location = new System.Drawing.Point(322, 175);
            this.lblAlineación.Name = "lblAlineación";
            this.lblAlineación.Size = new System.Drawing.Size(70, 16);
            this.lblAlineación.TabIndex = 14;
            this.lblAlineación.Text = "Alineación";
            // 
            // Libro1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1308, 546);
            this.Controls.Add(this.lblAlineación);
            this.Controls.Add(this.lblFuentes);
            this.Controls.Add(this.GBFuentes);
            this.Controls.Add(this.GBAlineación);
            this.Controls.Add(this.DGVCeldas);
            this.Name = "Libro1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Libro1";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Libro1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGVCeldas)).EndInit();
            this.GBAlineación.ResumeLayout(false);
            this.GBFuentes.ResumeLayout(false);
            this.GBFuentes.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView DGVCeldas;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column6;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column7;
        private System.Windows.Forms.Button btnNegrita;
        private System.Windows.Forms.Button btnCursiva;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column8;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column9;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column10;
        private System.Windows.Forms.ComboBox cmbFuentes;
        private System.Windows.Forms.ComboBox cmbTamaño;
        private System.Windows.Forms.Button btnSubrayado;
        private System.Windows.Forms.TextBox txtCeldaActiva;
        private System.Windows.Forms.GroupBox GBAlineación;
        private System.Windows.Forms.Button btnTXDerecha;
        private System.Windows.Forms.Button btnTXIzquierda;
        private System.Windows.Forms.Button btnTXCentro;
        private System.Windows.Forms.GroupBox GBFuentes;
        private System.Windows.Forms.Label lblFuentes;
        private System.Windows.Forms.Label lblAlineación;
        private System.Windows.Forms.Button btnTXAbajo;
        private System.Windows.Forms.Button btnTXArriba;
        private System.Windows.Forms.Button btnTXCentroV;
        private System.Windows.Forms.Button btnRellenoColor;
    }
}