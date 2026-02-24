using Kuni;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace Esclavitud
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tablas = this.Factory.CreateRibbonGroup();
            this.crear = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.button4 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.tablas.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.tablas);
            this.tab1.Label = "Esclavitud";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Label = "Inicio";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Contribuyentes";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.kuni);
            // 
            // tablas
            // 
            this.tablas.Items.Add(this.crear);
            this.tablas.Items.Add(this.button2);
            this.tablas.Items.Add(this.button3);
            this.tablas.Items.Add(this.separator1);
            this.tablas.Items.Add(this.button4);
            this.tablas.Label = "Tablas";
            this.tablas.Name = "tablas";
            // 
            // crear
            // 
            this.crear.Image = ((System.Drawing.Image)(resources.GetObject("crear.Image")));
            this.crear.Label = "Crear";
            this.crear.Name = "crear";
            this.crear.ShowImage = true;
            this.crear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.crearTablas);
            // 
            // button2
            // 
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "Importar";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.importar);
            // 
            // button3
            // 
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Label = "Guardar";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.Label = "Conciliar";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.tablas.ResumeLayout(false);
            this.tablas.PerformLayout();
            this.ResumeLayout(false);

        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup tablas;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton crear;

        public void kuni(object sender, RibbonControlEventArgs e)
        {
            System.Threading.Thread newWindowThread = new System.Threading.Thread(() =>
            {
                var wpfWindow = new Kuni.MainWindow();
                wpfWindow.Closed += (s, args) => System.Windows.Threading.Dispatcher.CurrentDispatcher.InvokeShutdown();
                wpfWindow.Show();
                System.Windows.Threading.Dispatcher.Run();
            });

            newWindowThread.SetApartmentState(System.Threading.ApartmentState.STA);
            newWindowThread.IsBackground = true;
            newWindowThread.Start();
        }

        public void crearTablas(object sender, RibbonControlEventArgs e)
        {
            new Class1().crearTablas();
        }

        public void importar(object sender, RibbonControlEventArgs e)
        {
            new Class1().importar();
        }

        internal RibbonGroup group1;
        internal RibbonButton button1;
        internal RibbonButton button2;
        internal RibbonButton button3;
        internal RibbonSeparator separator1;
        internal RibbonButton button4;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
