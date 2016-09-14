namespace Project_PERT_Add_in_2016
{
    partial class frmNiveisPERT
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
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.lblGeral = new System.Windows.Forms.Label();
            this.grpNiveis = new System.Windows.Forms.GroupBox();
            this.txtPessimista = new System.Windows.Forms.TextBox();
            this.txtEsperada = new System.Windows.Forms.TextBox();
            this.txtOtimista = new System.Windows.Forms.TextBox();
            this.lblPessimista = new System.Windows.Forms.Label();
            this.lblEsperada = new System.Windows.Forms.Label();
            this.lblOtimista = new System.Windows.Forms.Label();
            this.grpNiveis.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(13, 153);
            this.btnOK.Margin = new System.Windows.Forms.Padding(2);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(78, 21);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancelar
            // 
            this.btnCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancelar.Location = new System.Drawing.Point(111, 153);
            this.btnCancelar.Margin = new System.Windows.Forms.Padding(2);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(78, 21);
            this.btnCancelar.TabIndex = 1;
            this.btnCancelar.Text = "Cancel";
            this.btnCancelar.UseVisualStyleBackColor = true;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
            // 
            // lblGeral
            // 
            this.lblGeral.Location = new System.Drawing.Point(10, 7);
            this.lblGeral.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblGeral.Name = "lblGeral";
            this.lblGeral.Size = new System.Drawing.Size(179, 45);
            this.lblGeral.TabIndex = 2;
            this.lblGeral.Text = "Enter the levels of importance to the PERT calculations. These values must add 6." +
    "";
            // 
            // grpNiveis
            // 
            this.grpNiveis.Controls.Add(this.txtPessimista);
            this.grpNiveis.Controls.Add(this.txtEsperada);
            this.grpNiveis.Controls.Add(this.txtOtimista);
            this.grpNiveis.Controls.Add(this.lblPessimista);
            this.grpNiveis.Controls.Add(this.lblEsperada);
            this.grpNiveis.Controls.Add(this.lblOtimista);
            this.grpNiveis.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F);
            this.grpNiveis.Location = new System.Drawing.Point(12, 54);
            this.grpNiveis.Margin = new System.Windows.Forms.Padding(2);
            this.grpNiveis.Name = "grpNiveis";
            this.grpNiveis.Padding = new System.Windows.Forms.Padding(2);
            this.grpNiveis.Size = new System.Drawing.Size(177, 93);
            this.grpNiveis.TabIndex = 3;
            this.grpNiveis.TabStop = false;
            this.grpNiveis.Text = "Levels of importance";
            // 
            // txtPessimista
            // 
            this.txtPessimista.Location = new System.Drawing.Point(118, 63);
            this.txtPessimista.Margin = new System.Windows.Forms.Padding(2);
            this.txtPessimista.MaxLength = 1;
            this.txtPessimista.Name = "txtPessimista";
            this.txtPessimista.Size = new System.Drawing.Size(38, 19);
            this.txtPessimista.TabIndex = 5;
            this.txtPessimista.Text = "1";
            // 
            // txtEsperada
            // 
            this.txtEsperada.Location = new System.Drawing.Point(118, 41);
            this.txtEsperada.Margin = new System.Windows.Forms.Padding(2);
            this.txtEsperada.MaxLength = 1;
            this.txtEsperada.Name = "txtEsperada";
            this.txtEsperada.Size = new System.Drawing.Size(38, 19);
            this.txtEsperada.TabIndex = 4;
            this.txtEsperada.Text = "4";
            // 
            // txtOtimista
            // 
            this.txtOtimista.Location = new System.Drawing.Point(118, 18);
            this.txtOtimista.Margin = new System.Windows.Forms.Padding(2);
            this.txtOtimista.MaxLength = 1;
            this.txtOtimista.Name = "txtOtimista";
            this.txtOtimista.Size = new System.Drawing.Size(38, 19);
            this.txtOtimista.TabIndex = 3;
            this.txtOtimista.Text = "1";
            // 
            // lblPessimista
            // 
            this.lblPessimista.AutoSize = true;
            this.lblPessimista.Location = new System.Drawing.Point(5, 66);
            this.lblPessimista.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblPessimista.Name = "lblPessimista";
            this.lblPessimista.Size = new System.Drawing.Size(56, 13);
            this.lblPessimista.TabIndex = 2;
            this.lblPessimista.Text = "Pessimista";
            // 
            // lblEsperada
            // 
            this.lblEsperada.AutoSize = true;
            this.lblEsperada.Location = new System.Drawing.Point(5, 43);
            this.lblEsperada.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblEsperada.Name = "lblEsperada";
            this.lblEsperada.Size = new System.Drawing.Size(52, 13);
            this.lblEsperada.TabIndex = 1;
            this.lblEsperada.Text = "Esperada";
            // 
            // lblOtimista
            // 
            this.lblOtimista.AutoSize = true;
            this.lblOtimista.Location = new System.Drawing.Point(4, 20);
            this.lblOtimista.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblOtimista.Name = "lblOtimista";
            this.lblOtimista.Size = new System.Drawing.Size(52, 13);
            this.lblOtimista.TabIndex = 0;
            this.lblOtimista.Text = "Optimistic";
            // 
            // frmNiveisPERT
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.CancelButton = this.btnCancelar;
            this.ClientSize = new System.Drawing.Size(203, 184);
            this.Controls.Add(this.grpNiveis);
            this.Controls.Add(this.lblGeral);
            this.Controls.Add(this.btnCancelar);
            this.Controls.Add(this.btnOK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Location = new System.Drawing.Point(300, 300);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmNiveisPERT";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PERT Levels of importance";
            this.grpNiveis.ResumeLayout(false);
            this.grpNiveis.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Label lblGeral;
        private System.Windows.Forms.GroupBox grpNiveis;
        private System.Windows.Forms.TextBox txtPessimista;
        private System.Windows.Forms.TextBox txtEsperada;
        private System.Windows.Forms.TextBox txtOtimista;
        private System.Windows.Forms.Label lblPessimista;
        private System.Windows.Forms.Label lblEsperada;
        private System.Windows.Forms.Label lblOtimista;
    }
}