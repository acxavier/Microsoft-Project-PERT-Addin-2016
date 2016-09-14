using System;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;

namespace Project_PERT_Add_in_2016
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            string SKULanguage = @"Software\Microsoft\Office\16.0\MS Project";
            RegistryKey regKeyAppRoot = Registry.CurrentUser.OpenSubKey(SKULanguage, RegistryKeyPermissionCheck.ReadSubTree);
            SKULanguage = regKeyAppRoot.GetValue("CurrentUI").ToString();

            if ((SKULanguage == "1046") || (SKULanguage.Substring(0,2) == "pt"))
                Properties.Settings.Default.Language = "Portugues";
            else if ((SKULanguage == "1040") || (SKULanguage.Substring(0, 2) == "it"))
                Properties.Settings.Default.Language = "Italian";
            else if ((SKULanguage == "3082") || (SKULanguage.Substring(0, 2) == "it"))
                Properties.Settings.Default.Language = "Spanish";
            else if ((SKULanguage == "1036") || (SKULanguage.Substring(0, 2) == "fr"))
                Properties.Settings.Default.Language = "French";
            else
                Properties.Settings.Default.Language = "English";

            Ajustar();
        }

        
        private void btnCalcular_Click(object sender, RibbonControlEventArgs e)
        {
            if (SemProjetoAberto())
                return;
            
            Ajustar();

            frmCalcularPERT frmCalcular = new frmCalcularPERT();
            frmCalcular.ShowDialog();
        }

        private void btnOtimista_Click(object sender, RibbonControlEventArgs e)
        {
            if (SemProjetoAberto())
                return;
            
            Ajustar(); 
            
            if (Properties.Settings.Default.Language.ToString() == "English")
            {
                Globals.ThisAddIn.Application.ViewApply("Optimistic Gantt");
            }
            else if (Properties.Settings.Default.Language.ToString() == "Italian")
            {
                Globals.ThisAddIn.Application.ViewApply("Gantt ottimistico");
            }
            else if (Properties.Settings.Default.Language.ToString() == "Spanish")
            {
                Globals.ThisAddIn.Application.ViewApply("Gantt Optimista");
            }
            else if (Properties.Settings.Default.Language.ToString() == "French")
            {
                Globals.ThisAddIn.Application.ViewApply("Gantt Optimiste");
            }
            else
            {
                Globals.ThisAddIn.Application.ViewApply("Gantt otimista");
            }
        }

        private void btnEsperado_Click(object sender, RibbonControlEventArgs e)
        {
            if (SemProjetoAberto())
                return;
            
            Ajustar();

            if (Properties.Settings.Default.Language.ToString() == "English")
            {
                Globals.ThisAddIn.Application.ViewApply("Expected Gantt");
            }
            else if (Properties.Settings.Default.Language.ToString() == "Italian")
            {
                Globals.ThisAddIn.Application.ViewApply("Gantt prevista");
            }
            else if (Properties.Settings.Default.Language.ToString() == "Spanish")
            {
                Globals.ThisAddIn.Application.ViewApply("Gantt esperado");
            }
            else if (Properties.Settings.Default.Language.ToString() == "French")
            {
                Globals.ThisAddIn.Application.ViewApply("Gantt Probable");
            }
            else
            {
                Globals.ThisAddIn.Application.ViewApply("Gantt esperada");
            }
        }

        private void btnPessimista_Click(object sender, RibbonControlEventArgs e)
        {
            if (SemProjetoAberto())
                return;
            
            Ajustar();

            if (Properties.Settings.Default.Language.ToString() == "English")
            {
                Globals.ThisAddIn.Application.ViewApply("Pessimistic Gantt");
            }
            else if (Properties.Settings.Default.Language.ToString() == "Italian")
            {
                Globals.ThisAddIn.Application.ViewApply("Gantt pessimistico");
            }
            else if (Properties.Settings.Default.Language.ToString() == "Spanish")
            {
                Globals.ThisAddIn.Application.ViewApply("Gantt pesimista");
            }
            else if (Properties.Settings.Default.Language.ToString() == "French")
            {
                Globals.ThisAddIn.Application.ViewApply("Gantt Pessimiste");
            }
            else
            {
                Globals.ThisAddIn.Application.ViewApply("Gantt pessimista");
            }
        }

        private void btnFormularioEntrada_Click(object sender, RibbonControlEventArgs e)
        {
            if (SemProjetoAberto())
                return;
            
            Ajustar();

            if (Properties.Settings.Default.Language.ToString() == "English")
            {
                Globals.ThisAddIn.Application.ViewApply("Entry Sheet for PERT");
            }
            else if (Properties.Settings.Default.Language.ToString() == "Italian")
            {
                Globals.ThisAddIn.Application.ViewApply("Foglio inserimento per PERT");
            }
            else if (Properties.Settings.Default.Language.ToString() == "Spanish")
            {
                Globals.ThisAddIn.Application.ViewApply("Hoja de entradas PERT");
            }
            else if (Properties.Settings.Default.Language.ToString() == "French")
            {
                Globals.ThisAddIn.Application.ViewApply("Formulaire de saisie pour PERT");
            }
            else
            {
                Globals.ThisAddIn.Application.ViewApply("Planilha de entrada para PERT");
            }
        }

        private void btnNiveisPERT_Click(object sender, RibbonControlEventArgs e)
        {
            if (SemProjetoAberto())
                return;
            
            frmNiveisPERT frmNiveis = new frmNiveisPERT();
            frmNiveis.Show();
        }

        private Boolean SemProjetoAberto()
        {
            if (Globals.ThisAddIn.Application.ActiveProject == null)
            {
                string sMensagem = "Para executar o PERT Add-in, abra um projeto.";
                if (Properties.Settings.Default.Language.ToString() == "English")
                {
                    sMensagem = "To run the PERT add-in, open a project.";
                }
                else if (Properties.Settings.Default.Language.ToString() == "Italian")
                {
                    sMensagem = "Per eseguire il PERT Add-in, aprire un progetto.";
                }
                else if (Properties.Settings.Default.Language.ToString() == "Spanish")
                {
                    sMensagem = "Para ejecutar el PERT Add-in, abra un proyecto.";
                }
                else if (Properties.Settings.Default.Language.ToString() == "French")
                {
                    sMensagem = "Pour exécuter le PERT Add-in, ouvrez un projet.";
                }


                System.Windows.Forms.MessageBox.Show(sMensagem, "PERT Add-in", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return true;
            }
            else
                return false;
            
        }

        private void Ajustar()
        {
            object missing = System.Type.Missing;

            if (Properties.Settings.Default.Language.ToString() == "English")
            {
                btnOtimista.Label = "Optimistic";
                btnEsperado.Label = "Expected";
                btnPessimista.Label = "Pessimistic";
                btnFormularioEntrada.Label = "Entry Form";
                btnNiveisPERT.Label = "PERT Values";
                btnCalcular.Label = "Calculate";
                btnSobre.Label = "About";

                if (Globals.ThisAddIn.Application.ActiveProject == null)
                    return;

                // Criando tabela e visão para caso otimista
                Globals.ThisAddIn.Application.TableEdit("Optimistic table", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Optimistic table", true, missing, missing, missing, missing, "Indicators", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Optimistic table", true, missing, missing, missing, missing, "Name", "Task name", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Optimistic table", true, missing, missing, missing, missing, "Duration1", "Opt. Duration", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Optimistic table", true, missing, missing, missing, missing, "Start1", "Opt. Start", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Optimistic table", true, missing, missing, missing, missing, "Finish1", "Opt. Finish", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Optimistic Gantt", true, missing, 5, false, false, "Optimistic table", "All Tasks", "No Group");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Optimistic Gantt", false, missing, 5, false, false, "Optimistic table", "All Tasks", "No Group");
                }

                // Criando tabela e visão para caso esperada
                Globals.ThisAddIn.Application.TableEdit("Expected table", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Expected table", true, missing, missing, missing, missing, "Indicators", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Expected table", true, missing, missing, missing, missing, "Name", "Task name", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Expected table", true, missing, missing, missing, missing, "Duration2", "Exp. Duration", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Expected table", true, missing, missing, missing, missing, "Start2", "Exp. Start", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Expected table", true, missing, missing, missing, missing, "Finish2", "Exp. Finish", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Expected Gantt", true, missing, 5, false, false, "Expected table", "All Tasks", "No Group");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Expected Gantt", false, missing, 5, false, false, "Expected table", "All Tasks", "No Group");
                }

                // Criando tabela e visão para caso pessimista
                Globals.ThisAddIn.Application.TableEdit("Pessimistic table", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Pessimistic table", true, missing, missing, missing, missing, "Indicators", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Pessimistic table", true, missing, missing, missing, missing, "Name", "Task name", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Pessimistic table", true, missing, missing, missing, missing, "Duration3", "Pes. Duration", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Pessimistic table", true, missing, missing, missing, missing, "Start3", "Pes. Start", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Pessimistic table", true, missing, missing, missing, missing, "Finish3", "Pes. Finish", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Pessimistic Gantt", true, missing, 5, false, false, "Pessimistic table", "All Tasks", "No Group");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Pessimistic Gantt", false, missing, 5, false, false, "Pessimistic table", "All Tasks", "No Group");
                }

                // Criando tabela e visão para entrada de dados
                Globals.ThisAddIn.Application.TableEdit("Entry for PERT", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Entry for PERT", true, missing, missing, missing, missing, "Indicators", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Entry for PERT", true, missing, missing, missing, missing, "Name", "Task name", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Entry for PERT", true, missing, missing, missing, missing, "Duration", "Duration", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Entry for PERT", true, missing, missing, missing, missing, "Duration1", "Optimistic Dur.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Entry for PERT", true, missing, missing, missing, missing, "Duration2", "Expected Dur.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Entry for PERT", true, missing, missing, missing, missing, "Duration3", "Pessimistic Dur.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Entry Sheet for PERT", true, missing, 5, false, false, "Entry for PERT", "All Tasks", "No Group");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Entry Sheet for PERT", false, missing, 5, false, false, "Entry for PERT", "All Tasks", "No Group");
                }

            }
            else if (Properties.Settings.Default.Language.ToString() == "Italian")
            {
                btnOtimista.Label = "Ottimistico";
                btnEsperado.Label = "Prevista";
                btnPessimista.Label = "Pessimistico";
                btnFormularioEntrada.Label = "Immissione dati";
                btnNiveisPERT.Label = "Valori PERT";
                btnCalcular.Label = "Calcolo";
                btnSobre.Label = "Informazioni";

                if (Globals.ThisAddIn.Application.ActiveProject == null)
                    return;

                // Criando tabela e visão para caso otimista
                Globals.ThisAddIn.Application.TableEdit("Tabelle ottimistico", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle ottimistico", true, missing, missing, missing, missing, "Indicatori", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle ottimistico", true, missing, missing, missing, missing, "Nome", "Nome attività", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle ottimistico", true, missing, missing, missing, missing, "Durata1", "Dur. ottim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle ottimistico", true, missing, missing, missing, missing, "Inizio1", "Inizio ottim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle ottimistico", true, missing, missing, missing, missing, "Fine1", "Fine ottim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt ottimistico", true, missing, 5, false, false, "Tabelle ottimistico", "Tutte le attività", "Nessun raggruppamento");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt ottimistico", false, missing, 5, false, false, "Tabelle ottimistico", "Tutte le attività", "Nessun raggruppamento");
                }

                // Criando tabela e visão para Tabelle esperada
                Globals.ThisAddIn.Application.TableEdit("Tabelle prevista", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle prevista", true, missing, missing, missing, missing, "Indicatori", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle prevista", true, missing, missing, missing, missing, "Nome", "Nome attività", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle prevista", true, missing, missing, missing, missing, "Durata2", "Dur. prev.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle prevista", true, missing, missing, missing, missing, "Inizio2", "Inizio prev.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle prevista", true, missing, missing, missing, missing, "Fine2", "Fine prev.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt prevista", true, missing, 5, false, false, "Tabelle prevista", "Tutte le attività", "Nessun raggruppamento");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt prevista", false, missing, 5, false, false, "Tabelle prevista", "Tutte le attività", "Nessun raggruppamento");
                }

                // Criando tabela e visão para Tabelle pessimista
                Globals.ThisAddIn.Application.TableEdit("Tabelle pessimistico", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle pessimistico", true, missing, missing, missing, missing, "Indicatori", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle pessimistico", true, missing, missing, missing, missing, "Nome", "Nome attività", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle pessimistico", true, missing, missing, missing, missing, "Durata3", "Dur. pessim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle pessimistico", true, missing, missing, missing, missing, "Inizio3", "Inizio pessim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle pessimistico", true, missing, missing, missing, missing, "Fine3", "Fine pessim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt pessimistico", true, missing, 5, false, false, "Tabelle pessimistico", "Tutte le attività", "Nessun raggruppamento");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt pessimistico", false, missing, 5, false, false, "Tabelle pessimistico", "Tutte le attività", "Nessun raggruppamento");
                }

                // Criando tabela e visão para entrada de dados
                Globals.ThisAddIn.Application.TableEdit("Tabelle inserimento per PERT AP", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle inserimento per PERT AP", true, missing, missing, missing, missing, "Indicatori", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle inserimento per PERT AP", true, missing, missing, missing, missing, "Nome", "Nome attività", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle inserimento per PERT AP", true, missing, missing, missing, missing, "Durata", "Durata1", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle inserimento per PERT AP", true, missing, missing, missing, missing, "Durata1", "Dur. ottim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle inserimento per PERT AP", true, missing, missing, missing, missing, "Durata2", "Dur. prevista", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabelle inserimento per PERT AP", true, missing, missing, missing, missing, "Durata3", "Dur. pessim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Foglio inserimento per PERT", true, missing, 5, false, false, "Tabelle inserimento per PERT AP", "Tutte le attività", "Nessun raggruppamento");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Foglio inserimento per PERT", false, missing, 5, false, false, "Tabelle inserimento per PERT AP", "Tutte le attività", "Nessun raggruppamento");
                }
            }
            else if (Properties.Settings.Default.Language.ToString() == "Spanish")
            {
                btnOtimista.Label = "Optimista";
                btnEsperado.Label = "Esperada";
                btnPessimista.Label = "Pesimista";
                btnFormularioEntrada.Label = "Hoja entradas";
                btnNiveisPERT.Label = "Valores PERT";
                btnCalcular.Label = "Calcular";
                btnSobre.Label = "Sobre";

                if (Globals.ThisAddIn.Application.ActiveProject == null)
                    return;

                // Criando tabela e visão para caso otimista
                Globals.ThisAddIn.Application.TableEdit("Tabla optimista", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla optimista", true, missing, missing, missing, missing, "Indicadores", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla optimista", true, missing, missing, missing, missing, "Nombre", "Nombre da tarea", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla optimista", true, missing, missing, missing, missing, "Duración1", "Dur. optim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla optimista", true, missing, missing, missing, missing, "Comienzo1", "Com. optim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla optimista", true, missing, missing, missing, missing, "Fin1", "Fin optim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt optimista", true, missing, 5, false, false, "Tabla optimista", "Todas las tareas", "Sin agrupar");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt optimista", false, missing, 5, false, false, "Tabla optimista", "Todas las tareas", "Sin agrupar");
                }

                // Criando tabela e visão para caso esperada
                Globals.ThisAddIn.Application.TableEdit("Tabla esperada", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla esperada", true, missing, missing, missing, missing, "Indicadores", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla esperada", true, missing, missing, missing, missing, "Nombre", "Nombre da tarea", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla esperada", true, missing, missing, missing, missing, "Duración2", "Dur. espe.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla esperada", true, missing, missing, missing, missing, "Comienzo2", "Com. espe.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla esperada", true, missing, missing, missing, missing, "Fin2", "Fin espe.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt esperado", true, missing, 5, false, false, "Tabla esperada", "Todas las tareas", "Sin agrupar");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt esperado", false, missing, 5, false, false, "Tabla esperada", "Todas las tareas", "Sin agrupar");
                }

                // Criando tabela e visão para caso pessimista
                Globals.ThisAddIn.Application.TableEdit("Tabla pesimista", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla pesimista", true, missing, missing, missing, missing, "Indicadores", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla pesimista", true, missing, missing, missing, missing, "Nombre", "Nombre da tarea", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla pesimista", true, missing, missing, missing, missing, "Duración3", "Dur. pes.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla pesimista", true, missing, missing, missing, missing, "Comienzo3", "Com. pes.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla pesimista", true, missing, missing, missing, missing, "Fin3", "Fin pes.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt pesimista", true, missing, 5, false, false, "Tabla pesimista", "Todas las tareas", "Sin agrupar");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt pesimista", false, missing, 5, false, false, "Tabla pesimista", "Todas las tareas", "Sin agrupar");
                }

                // Criando tabela e visão para entrada de dados
                Globals.ThisAddIn.Application.TableEdit("Tabla de entrada PERT", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla de entrada PERT", true, missing, missing, missing, missing, "Indicadores", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla de entrada PERT", true, missing, missing, missing, missing, "Nombre", "Nombre da tarea", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla de entrada PERT", true, missing, missing, missing, missing, "Duración", "Duración", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla de entrada PERT", true, missing, missing, missing, missing, "Duración1", "Dur. Optimista", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla de entrada PERT", true, missing, missing, missing, missing, "Duración2", "Dur. Esperada", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Tabla de entrada PERT", true, missing, missing, missing, missing, "Duración3", "Dur. Pesimista", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Hoja de entradas PERT", true, missing, 5, false, false, "Tabla de entrada PERT", "Todas las tareas", "Sin agrupar");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Hoja de entradas PERT", false, missing, 5, false, false, "Tabla de entrada PERT", "Todas las tareas", "Sin agrupar");
                }

            }
            else if (Properties.Settings.Default.Language.ToString() == "French")
            {
                btnOtimista.Label = "Optimiste";
                btnEsperado.Label = "Probable";
                btnPessimista.Label = "Pessimiste";
                btnFormularioEntrada.Label = "Formulaire de saisie";
                btnNiveisPERT.Label = "Valeurs PERT";
                btnCalcular.Label = "Calculer";
                btnSobre.Label = "Sur";

                if (Globals.ThisAddIn.Application.ActiveProject == null)
                    return;

                // Criando tabela e visão para caso otimista
                Globals.ThisAddIn.Application.TableEdit("Table Optimiste", true, true, true, missing, "N°", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Optimiste", true, missing, missing, missing, missing, "Indicateurs", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Optimiste", true, missing, missing, missing, missing, "Nom", "Nom de la tâche", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Optimiste", true, missing, missing, missing, missing, "Durée1", "Dur. optim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Optimiste", true, missing, missing, missing, missing, "Début1", "Début optim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Optimiste", true, missing, missing, missing, missing, "Fin1", "Fin optim.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt Optimiste", true, missing, 5, false, false, "Table Optimiste", "Toutes les tâches", "Aucun groupe");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt Optimiste", false, missing, 5, false, false, "Table Optimiste", "Toutes les tâches", "Aucun groupe");
                }

                // Criando tabela e visão para caso esperada
                Globals.ThisAddIn.Application.TableEdit("Table Probable", true, true, true, missing, "N°", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Probable", true, missing, missing, missing, missing, "Indicateurs", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Probable", true, missing, missing, missing, missing, "Nom", "Nom de la tâche", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Probable", true, missing, missing, missing, missing, "Durée2", "Dur. Prob.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Probable", true, missing, missing, missing, missing, "Début2", "Début Prob.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Probable", true, missing, missing, missing, missing, "Fin2", "Fin Prob.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt Probable", true, missing, 5, false, false, "Table Probable", "Toutes les tâches", "Aucun groupe");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt Probable", false, missing, 5, false, false, "Table Probable", "Toutes les tâches", "Aucun groupe");
                }

                // Criando tabela e visão para caso pessimista
                Globals.ThisAddIn.Application.TableEdit("Table Pessimiste", true, true, true, missing, "N°", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Pessimiste", true, missing, missing, missing, missing, "Indicateurs", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Pessimiste", true, missing, missing, missing, missing, "Nom", "Nom de la tâche", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Pessimiste", true, missing, missing, missing, missing, "Durée3", "Dur. pes.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Pessimiste", true, missing, missing, missing, missing, "Début3", "Début pes.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table Pessimiste", true, missing, missing, missing, missing, "Fin3", "Fin pes.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt Pessimiste", true, missing, 5, false, false, "Table Pessimiste", "Toutes les tâches", "Aucun groupe");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt Pessimiste", false, missing, 5, false, false, "Table Pessimiste", "Toutes les tâches", "Aucun groupe");
                }

                // Criando tabela e visão para entrada de dados
                Globals.ThisAddIn.Application.TableEdit("Table de saisie PERT", true, true, true, missing, "N°", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table de saisie PERT", true, missing, missing, missing, missing, "Indicateurs", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table de saisie PERT", true, missing, missing, missing, missing, "Nom", "Nom de la tâche", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table de saisie PERT", true, missing, missing, missing, missing, "Durée", "Durée", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table de saisie PERT", true, missing, missing, missing, missing, "Durée1", "Dur. Optimiste", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table de saisie PERT", true, missing, missing, missing, missing, "Durée2", "Dur. Probable", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Table de saisie PERT", true, missing, missing, missing, missing, "Durée3", "Dur. Pessimiste", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Formulaire de saisie pour PERT", true, missing, 5, false, false, "Table de saisie PERT", "Toutes les tâches", "Aucun groupe");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Formulaire de saisie pour PERT", false, missing, 5, false, false, "Table de saisie PERT", "Toutes les tâches", "Aucun groupe");
                }

            }
            else
            {
                btnOtimista.Label = "Otimista";
                btnEsperado.Label = "Esperada";
                btnPessimista.Label = "Pessimista";
                btnFormularioEntrada.Label = "Formulário de Entrada";
                btnNiveisPERT.Label = "Peso PERT";
                btnCalcular.Label = "Calcular";
                btnSobre.Label = "Sobre";

                if (Globals.ThisAddIn.Application.ActiveProject == null)
                    return;

                // Criando tabela e visão para caso otimista
                Globals.ThisAddIn.Application.TableEdit("Caso otimista", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso otimista", true, missing, missing, missing, missing, "Indicadores", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso otimista", true, missing, missing, missing, missing, "Nome", "Nome da tarefa", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso otimista", true, missing, missing, missing, missing, "Duração1", "Dur otim", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso otimista", true, missing, missing, missing, missing, "Início1", "Início otim", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso otimista", true, missing, missing, missing, missing, "Término1", "Término otim", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt otimista", true, missing, 5, false, false, "Caso otimista", "Todas as tarefas", "Nenhum grupo");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt otimista", false, missing, 5, false, false, "Caso otimista", "Todas as tarefas", "Nenhum grupo");
                }

                // Criando tabela e visão para caso esperada
                Globals.ThisAddIn.Application.TableEdit("Caso esperada", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso esperada", true, missing, missing, missing, missing, "Indicadores", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso esperada", true, missing, missing, missing, missing, "Nome", "Nome da tarefa", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso esperada", true, missing, missing, missing, missing, "Duração2", "Dur esp.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso esperada", true, missing, missing, missing, missing, "Início2", "Início esp.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso esperada", true, missing, missing, missing, missing, "Término2", "Término esp.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt esperada", true, missing, 5, false, false, "Caso esperada", "Todas as tarefas", "Nenhum grupo");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt esperada", false, missing, 5, false, false, "Caso esperada", "Todas as tarefas", "Nenhum grupo");
                }

                // Criando tabela e visão para caso pessimista
                Globals.ThisAddIn.Application.TableEdit("Caso pessimista", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso pessimista", true, missing, missing, missing, missing, "Indicadores", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso pessimista", true, missing, missing, missing, missing, "Nome", "Nome da tarefa", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso pessimista", true, missing, missing, missing, missing, "Duração3", "Dur pes.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso pessimista", true, missing, missing, missing, missing, "Início3", "Início pes.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Caso pessimista", true, missing, missing, missing, missing, "Término3", "Término pes.", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt pessimista", true, missing, 5, false, false, "Caso pessimista", "Todas as tarefas", "Nenhum grupo");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Gantt pessimista", false, missing, 5, false, false, "Caso pessimista", "Todas as tarefas", "Nenhum grupo");
                }

                // Criando tabela e visão para entrada de dados
                Globals.ThisAddIn.Application.TableEdit("Entrada PERT", true, true, true, missing, "Id", missing, "", 5, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Entrada PERT", true, missing, missing, missing, missing, "Indicadores", "", 6, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Entrada PERT", true, missing, missing, missing, missing, "Nome", "Nome da tarefa", 75, 0, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Entrada PERT", true, missing, missing, missing, missing, "Duração", "Duração", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Entrada PERT", true, missing, missing, missing, missing, "Duração1", "Dur. otimista", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Entrada PERT", true, missing, missing, missing, missing, "Duração2", "Dur. esperada", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                Globals.ThisAddIn.Application.TableEdit("Entrada PERT", true, missing, missing, missing, missing, "Duração3", "Dur. pessimista", 19, 1, false, true, 255, 1, missing, 1, missing, missing);
                try
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Planilha de entrada para PERT", true, missing, 5, false, false, "Entrada PERT", "Todas as tarefas", "Nenhum grupo");
                }
                catch
                {
                    Globals.ThisAddIn.Application.ViewEditSingle("Planilha de entrada para PERT", false, missing, 5, false, false, "Entrada PERT", "Todas as tarefas", "Nenhum grupo");
                }
            }

        }

        private void btnSobre_Click(object sender, RibbonControlEventArgs e)
        {
            frmSobre frmSobrePERT = new frmSobre();
            frmSobrePERT.ShowDialog();
        }
    }
}
