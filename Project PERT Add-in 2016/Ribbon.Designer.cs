namespace Project_PERT_Add_in_2016
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabBHS = this.Factory.CreateRibbonTab();
            this.grpPERT = this.Factory.CreateRibbonGroup();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.btnOtimista = this.Factory.CreateRibbonButton();
            this.buttonGroup3 = this.Factory.CreateRibbonButtonGroup();
            this.btnEsperado = this.Factory.CreateRibbonButton();
            this.buttonGroup2 = this.Factory.CreateRibbonButtonGroup();
            this.btnPessimista = this.Factory.CreateRibbonButton();
            this.btnFormularioEntrada = this.Factory.CreateRibbonButton();
            this.btnNiveisPERT = this.Factory.CreateRibbonButton();
            this.btnCalcular = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnSobre = this.Factory.CreateRibbonButton();
            this.tabBHS.SuspendLayout();
            this.grpPERT.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.buttonGroup3.SuspendLayout();
            this.buttonGroup2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabBHS
            // 
            this.tabBHS.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabBHS.Groups.Add(this.grpPERT);
            this.tabBHS.Label = "PERT Add-in";
            this.tabBHS.Name = "tabBHS";
            // 
            // grpPERT
            // 
            this.grpPERT.Items.Add(this.buttonGroup1);
            this.grpPERT.Items.Add(this.buttonGroup3);
            this.grpPERT.Items.Add(this.buttonGroup2);
            this.grpPERT.Items.Add(this.btnFormularioEntrada);
            this.grpPERT.Items.Add(this.btnNiveisPERT);
            this.grpPERT.Items.Add(this.btnCalcular);
            this.grpPERT.Items.Add(this.separator1);
            this.grpPERT.Items.Add(this.btnSobre);
            this.grpPERT.Label = "PERT";
            this.grpPERT.Name = "grpPERT";
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.btnOtimista);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // btnOtimista
            // 
            this.btnOtimista.Image = global::Project_PERT_Add_in_2016.Properties.Resources.Otimista;
            this.btnOtimista.Label = "Optimistic";
            this.btnOtimista.Name = "btnOtimista";
            this.btnOtimista.ScreenTip = "Gantt Otimista";
            this.btnOtimista.ShowImage = true;
            this.btnOtimista.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOtimista_Click);
            // 
            // buttonGroup3
            // 
            this.buttonGroup3.Items.Add(this.btnEsperado);
            this.buttonGroup3.Name = "buttonGroup3";
            // 
            // btnEsperado
            // 
            this.btnEsperado.Image = global::Project_PERT_Add_in_2016.Properties.Resources.Esperado;
            this.btnEsperado.Label = "Expected";
            this.btnEsperado.Name = "btnEsperado";
            this.btnEsperado.ScreenTip = "Gantt Esperada";
            this.btnEsperado.ShowImage = true;
            this.btnEsperado.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEsperado_Click);
            // 
            // buttonGroup2
            // 
            this.buttonGroup2.Items.Add(this.btnPessimista);
            this.buttonGroup2.Name = "buttonGroup2";
            // 
            // btnPessimista
            // 
            this.btnPessimista.Image = global::Project_PERT_Add_in_2016.Properties.Resources.Pessimista;
            this.btnPessimista.Label = "Pessimistic";
            this.btnPessimista.Name = "btnPessimista";
            this.btnPessimista.ScreenTip = "Gantt Pessimista";
            this.btnPessimista.ShowImage = true;
            this.btnPessimista.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPessimista_Click);
            // 
            // btnFormularioEntrada
            // 
            this.btnFormularioEntrada.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFormularioEntrada.Label = "Entry Form";
            this.btnFormularioEntrada.Name = "btnFormularioEntrada";
            this.btnFormularioEntrada.OfficeImageId = "Consolidate";
            this.btnFormularioEntrada.ScreenTip = "Formulário de entrada PERT";
            this.btnFormularioEntrada.ShowImage = true;
            this.btnFormularioEntrada.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormularioEntrada_Click);
            // 
            // btnNiveisPERT
            // 
            this.btnNiveisPERT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNiveisPERT.Label = "PERT Values";
            this.btnNiveisPERT.Name = "btnNiveisPERT";
            this.btnNiveisPERT.OfficeImageId = "EditBusinessDataEntity";
            this.btnNiveisPERT.ScreenTip = "Definir níveis de importância PERT";
            this.btnNiveisPERT.ShowImage = true;
            this.btnNiveisPERT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNiveisPERT_Click);
            // 
            // btnCalcular
            // 
            this.btnCalcular.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCalcular.Label = "Calculate";
            this.btnCalcular.Name = "btnCalcular";
            this.btnCalcular.OfficeImageId = "CalculationOptionsMenu";
            this.btnCalcular.ScreenTip = "Calcular PERT";
            this.btnCalcular.ShowImage = true;
            this.btnCalcular.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalcular_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnSobre
            // 
            this.btnSobre.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSobre.Label = "About";
            this.btnSobre.Name = "btnSobre";
            this.btnSobre.OfficeImageId = "FindUnread";
            this.btnSobre.ShowImage = true;
            this.btnSobre.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSobre_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Project.Project";
            this.Tabs.Add(this.tabBHS);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tabBHS.ResumeLayout(false);
            this.tabBHS.PerformLayout();
            this.grpPERT.ResumeLayout(false);
            this.grpPERT.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.buttonGroup3.ResumeLayout(false);
            this.buttonGroup3.PerformLayout();
            this.buttonGroup2.ResumeLayout(false);
            this.buttonGroup2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabBHS;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpPERT;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalcular;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPessimista;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEsperado;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOtimista;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormularioEntrada;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNiveisPERT;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSobre;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
