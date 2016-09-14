namespace Project_PERT_Add_in_2016
{
    partial class frmCalcularPERT
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCalcularPERT));
            this.btnNao = new System.Windows.Forms.Button();
            this.btnSim = new System.Windows.Forms.Button();
            this.lblGeral1 = new System.Windows.Forms.Label();
            this.lblGeral2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnNao
            // 
            this.btnNao.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnNao.Location = new System.Drawing.Point(267, 98);
            this.btnNao.Margin = new System.Windows.Forms.Padding(2);
            this.btnNao.Name = "btnNao";
            this.btnNao.Size = new System.Drawing.Size(78, 21);
            this.btnNao.TabIndex = 3;
            this.btnNao.Text = "No";
            this.btnNao.UseVisualStyleBackColor = true;
            this.btnNao.Click += new System.EventHandler(this.btnNao_Click);
            // 
            // btnSim
            // 
            this.btnSim.Location = new System.Drawing.Point(177, 98);
            this.btnSim.Margin = new System.Windows.Forms.Padding(2);
            this.btnSim.Name = "btnSim";
            this.btnSim.Size = new System.Drawing.Size(78, 21);
            this.btnSim.TabIndex = 2;
            this.btnSim.Text = "Yes";
            this.btnSim.UseVisualStyleBackColor = true;
            this.btnSim.Click += new System.EventHandler(this.btnSim_Click);
            // 
            // lblGeral1
            // 
            this.lblGeral1.Location = new System.Drawing.Point(9, 7);
            this.lblGeral1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblGeral1.Name = "lblGeral1";
            this.lblGeral1.Size = new System.Drawing.Size(485, 57);
            this.lblGeral1.TabIndex = 4;
            this.lblGeral1.Text = resources.GetString("lblGeral1.Text");
            // 
            // lblGeral2
            // 
            this.lblGeral2.AutoSize = true;
            this.lblGeral2.Location = new System.Drawing.Point(11, 67);
            this.lblGeral2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblGeral2.Name = "lblGeral2";
            this.lblGeral2.Size = new System.Drawing.Size(95, 13);
            this.lblGeral2.TabIndex = 5;
            this.lblGeral2.Text = "Want to continue?";
            // 
            // frmCalcularPERT
            // 
            this.AcceptButton = this.btnSim;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.CancelButton = this.btnNao;
            this.ClientSize = new System.Drawing.Size(497, 129);
            this.Controls.Add(this.lblGeral2);
            this.Controls.Add(this.lblGeral1);
            this.Controls.Add(this.btnNao);
            this.Controls.Add(this.btnSim);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Location = new System.Drawing.Point(300, 300);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmCalcularPERT";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PERT Analysis";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnNao;
        private System.Windows.Forms.Button btnSim;
        private System.Windows.Forms.Label lblGeral1;
        private System.Windows.Forms.Label lblGeral2;
    }
}