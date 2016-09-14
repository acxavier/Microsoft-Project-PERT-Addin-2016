using System;
using System.Windows.Forms;
using MSProject = Microsoft.Office.Interop.MSProject;


namespace Project_PERT_Add_in_2016
{
    public partial class frmCalcularPERT : Form
    {
        
        public frmCalcularPERT()
        {
            InitializeComponent();

            if (Properties.Settings.Default.Language.ToString() == "English")
            {
                this.Text = "PERT Analysis";
                lblGeral1.Text = "The PERT analysis uses the information contained in \"Dur. optimistic\", \"Dur. expected\"and \"Dur. pessimistic\". These values are stored in the Duration1, Duration2 and Duration3 respectively. Click \"Yes\" to recalculate the field \"Duration\" based on these three fields.";
                lblGeral2.Text = "Want to continue?";
                btnSim.Text = "Yes";
                btnNao.Text = "No";
            }
            else if (Properties.Settings.Default.Language.ToString() == "Italian")
            {
                this.Text = "Analisi PERT";
                lblGeral1.Text = "L' analisi PERT utilizza le informazioni contenute nei campi \"Dur. ottimistica\", \"Dur. prevista\", e \"Durata pessimistica\". Questi valori vengono memorizzati rispettivamente nei campi personalizzati \"Durata 1\", \"Durata 2\" e \"Durata 3\". Clicca \"Si\" per ricalcolare il campo \"Durata\", basandolo sui valori di questi tre campi.";
                lblGeral2.Text = "Continuare ?";
                btnSim.Text = "Si";
                btnNao.Text = "No";
            }
            else if (Properties.Settings.Default.Language.ToString() == "Spanish")
            {
                this.Text = "Análisis PERT";
                lblGeral1.Text = "El Análisis PERT utiliza la información contenida en los campos \"Dur. optimista\", \"Dur. esperada\" y \"Dur. pesimista\". Los valores se guardan en los campos Duración1, Duración2 y Duración3 respectivamente. Al hacer clic en Sí re recalcula el valor del campo Duración tomando base los de estos tres campos.";
                lblGeral2.Text = "¿Desea continuar?";
                btnSim.Text = "Sí";
                btnNao.Text = "No";
            }
            else if (Properties.Settings.Default.Language.ToString() == "French")
            {
                this.Text = "Analyse PERT";
                lblGeral1.Text = "L'analyse PERT utilise les informations contenues dans \"Dur. Optimiste\", \"Durée Prob.\" et \"Durée Pes.\". Ces valeurs sont stockées respectivement dans Durée1, Durée2 et Durée3. Cliquer sur \"Oui\" pour recalculer les champs \"Durée\" à partir de ces 3 champs.";
                lblGeral2.Text = "Voulez-vous continuer?";
                btnSim.Text = "Oui";
                btnNao.Text = "Non";
            }
            else
            {
                this.Text = "Análise PERT";
                lblGeral1.Text = "A análise PERT usa as informações contidas em \"Dur. otimista\", \"Dur. esperada\" e \"Dur. pessimista\". Esses valores são armazenados nos campos Duração1, Duração2 e Duração3 respectivamente. Clique em \"Sim\" para recalcular o campo \"Duração\" baseado nesses três campos.";
                lblGeral2.Text = "Deseja continuar?";
                btnSim.Text = "Sim";
                btnNao.Text = "Não";
            }
        }

        private void btnSim_Click(object sender, EventArgs e)
        {
            MSProject.Project proj = Globals.ThisAddIn.Application.ActiveProject;

            try
            {
                
                double iOtimista = proj.ProjectSummaryTask.Number1;
                double iEsperada = proj.ProjectSummaryTask.Number2;
                double iPessimista = proj.ProjectSummaryTask.Number3;

                if (iOtimista + iEsperada + iPessimista != 6)
                {
                    iOtimista = 1;
                    iEsperada = 4;
                    iPessimista = 1;
                }

                foreach (MSProject.Task tskT in proj.Tasks)
                {
                    if ((tskT.PercentComplete == 0) & (tskT.PercentWorkComplete == 0) & (tskT.Duration1 + tskT.Duration2 + tskT.Duration3 != 0))
                    {
                        tskT.Duration = ((((tskT.Duration1) * iOtimista) + ((tskT.Duration2) * iEsperada) + ((tskT.Duration3) * iPessimista)) / 6);
                        
                    }

                }
            }
            catch (System.Exception Ex)
            {
                MessageBox.Show(Ex.Message);
            }

            this.Close();
            
        }

        private void btnNao_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
