using System;
using System.Windows.Forms;

namespace Project_PERT_Add_in_2016
{
    public partial class frmSobre : Form
    {
        public frmSobre()
        {
            InitializeComponent();
            string sVersion = "1.0";

            if (Properties.Settings.Default.Language.ToString() == "English")
            {
                this.Text = "About";
                okButton.Text = "Close";
                lblVersao.Text = "Version: " + sVersion;
            }
            else if (Properties.Settings.Default.Language.ToString() == "Italian")
            {
                this.Text = "Informazioni";
                okButton.Text = "Chiudi";
                lblVersao.Text = "Versione: " + sVersion;
            }
            else if (Properties.Settings.Default.Language.ToString() == "Spanish")
            {
                this.Text = "Sobre";
                okButton.Text = "Cerrar";
                lblVersao.Text = "Versión: " + sVersion;
            }
            else if (Properties.Settings.Default.Language.ToString() == "French")
            {
                this.Text = "A propos";
                okButton.Text = "Fermer";
                lblVersao.Text = "Version: " + sVersion;
            }
            else
            {
                this.Text = "Sobre";
                okButton.Text = "Fechar";
                lblVersao.Text = "Versão: " + sVersion;
            }
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
