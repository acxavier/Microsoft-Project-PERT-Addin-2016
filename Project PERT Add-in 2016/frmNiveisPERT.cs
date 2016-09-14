using System;
using System.Windows.Forms;
using MSProject = Microsoft.Office.Interop.MSProject;


namespace Project_PERT_Add_in_2016
{
    public partial class frmNiveisPERT : Form
    {
        public MSProject.Project proj = Globals.ThisAddIn.Application.ActiveProject;
        public frmNiveisPERT()
        {
            InitializeComponent();

            if (proj.ProjectSummaryTask.Number1 + proj.ProjectSummaryTask.Number2 + proj.ProjectSummaryTask.Number3 == 6)
            {
                txtOtimista.Text = proj.ProjectSummaryTask.Number1.ToString();
                txtEsperada.Text = proj.ProjectSummaryTask.Number2.ToString();
                txtPessimista.Text = proj.ProjectSummaryTask.Number3.ToString();
            }
            else
            {
                txtOtimista.Text = "1";
                txtEsperada.Text = "4";
                txtPessimista.Text = "1";
            }

            if (Properties.Settings.Default.Language.ToString() == "English")
            {
                this.Text = "PERT Levels of importance";
                lblGeral.Text = "Enter the levels of importance to the PERT calculations. These values must add 6.";
                grpNiveis.Text = "Levels of importance";
                lblOtimista.Text = "Optimistic";
                lblEsperada.Text = "Expected";
                lblPessimista.Text = "Pessimistic";
                btnCancelar.Text = "Cancel";
            }
            else if (Properties.Settings.Default.Language.ToString() == "Italian")
            {
                this.Text = "Pesi delle stime di durata PERT";
                lblGeral.Text = "Inserisci il peso delle stime di durata per i calcoli PERT. La somma dei valori deve essere 6.";
                grpNiveis.Text = "Pesi delle stime di durata";
                lblOtimista.Text = "Ottimistico";
                lblEsperada.Text = "Previsto";
                lblPessimista.Text = "Pessimistico";
                btnCancelar.Text = "Cancellare";
            }
            else if (Properties.Settings.Default.Language.ToString() == "Spanish")
            {
                this.Text = "Establecer pesos PERT";
                lblGeral.Text = "Escriba los pesos para realizar los cálculos PERT. La suma de los valores debe ser 6.";
                grpNiveis.Text = "Pesos de duración";
                lblOtimista.Text = "Optimista";
                lblEsperada.Text = "Esperada";
                lblPessimista.Text = "Pesimista";
                btnCancelar.Text = "Cancelar";
            }
            else if (Properties.Settings.Default.Language.ToString() == "French")
            {
                this.Text = "Pondérations PERT";
                lblGeral.Text = "Entrer les pondérations de calcul PERT. Les valeurs doivent ajouter 6.";
                grpNiveis.Text = "Pondérations";
                lblOtimista.Text = "Optimiste";
                lblEsperada.Text = "Probable";
                lblPessimista.Text = "Pessimiste";
                btnCancelar.Text = "Annuler";
            }
            else
            {
                this.Text = "Níveis  de importância PERT";
                lblGeral.Text = "Insira os níveis de importância para os cálculos PERT. Esses valores devem somar 6.";
                grpNiveis.Text = "Níveis de importância";
                lblOtimista.Text = "Otimista";
                lblEsperada.Text = "Esperada";
                lblPessimista.Text = "Pessimista";
                btnCancelar.Text = "Cancelar";
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (Convert.ToDouble(txtOtimista.Text) + Convert.ToDouble(txtEsperada.Text) + Convert.ToDouble(txtPessimista.Text) == 6)
            {
                proj.ProjectSummaryTask.Number1 = Convert.ToDouble(txtOtimista.Text);
                proj.ProjectSummaryTask.Number2 = Convert.ToDouble(txtEsperada.Text);
                proj.ProjectSummaryTask.Number3 = Convert.ToDouble(txtPessimista.Text);

                this.Close();
            }
            else
            {
                if (Properties.Settings.Default.Language.ToString() == "English")
                    MessageBox.Show("The values must add 6.", "Microsoft Project", MessageBoxButtons.OK);
                else if (Properties.Settings.Default.Language.ToString() == "Italian")
                    MessageBox.Show("La somma dei valori deve essere pari a 6.", "Microsoft Project", MessageBoxButtons.OK);
                else if (Properties.Settings.Default.Language.ToString() == "Spanish")
                    MessageBox.Show("La suma de los valores debe ser 6.", "Microsoft Project", MessageBoxButtons.OK);
                else if (Properties.Settings.Default.Language.ToString() == "French")
                    MessageBox.Show("Les valeurs doivent ajouter 6.", "Microsoft Project", MessageBoxButtons.OK);
                else
                    MessageBox.Show("Os valores dos níveis de importância devem somar 6.", "Microsoft Project", MessageBoxButtons.OK);
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
