namespace WordSelectionTool
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.tabSelectTools = this.Factory.CreateRibbonTab();
            this.groupSelectTables = this.Factory.CreateRibbonGroup();
            this.btnSelectAllTables = this.Factory.CreateRibbonButton();
            this.btnSelectFirstRow = this.Factory.CreateRibbonButton();
            this.btnSelectLastRow = this.Factory.CreateRibbonButton();
            this.btnSelectFirstColumn = this.Factory.CreateRibbonButton();
            this.btnSelectLastColumn = this.Factory.CreateRibbonButton();
            this.btnSelect1RowTables = this.Factory.CreateRibbonButton();
            this.btnSelectMultiRowTables = this.Factory.CreateRibbonButton();
            this.btnSelectFirstCellInTables = this.Factory.CreateRibbonButton();
            this.groupSelectLists = this.Factory.CreateRibbonGroup();
            this.btnSelectAllLists = this.Factory.CreateRibbonButton();
            this.btnSelectNumericLists = this.Factory.CreateRibbonButton();
            this.btnSelectBulletLists = this.Factory.CreateRibbonButton();
            this.groupSelectObjects = this.Factory.CreateRibbonGroup();
            this.btnSelectCurrentPage = this.Factory.CreateRibbonButton();
            this.btnSelectFormulas = this.Factory.CreateRibbonButton();
            this.btnSelectShapes = this.Factory.CreateRibbonButton();
            this.groupHelp = this.Factory.CreateRibbonGroup();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabSelectTools.SuspendLayout();
            this.groupSelectTables.SuspendLayout();
            this.groupSelectLists.SuspendLayout();
            this.groupSelectObjects.SuspendLayout();
            this.groupHelp.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabSelectTools
            // 
            this.tabSelectTools.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabSelectTools.Groups.Add(this.groupSelectTables);
            this.tabSelectTools.Groups.Add(this.groupSelectLists);
            this.tabSelectTools.Groups.Add(this.groupSelectObjects);
            this.tabSelectTools.Groups.Add(this.groupHelp);
            this.tabSelectTools.Label = "Selection Tool";
            this.tabSelectTools.Name = "tabSelectTools";
            // 
            // groupSelectTables
            // 
            this.groupSelectTables.Items.Add(this.btnSelectAllTables);
            this.groupSelectTables.Items.Add(this.btnSelectFirstRow);
            this.groupSelectTables.Items.Add(this.btnSelectLastRow);
            this.groupSelectTables.Items.Add(this.btnSelectFirstColumn);
            this.groupSelectTables.Items.Add(this.btnSelectLastColumn);
            this.groupSelectTables.Items.Add(this.btnSelect1RowTables);
            this.groupSelectTables.Items.Add(this.btnSelectMultiRowTables);
            this.groupSelectTables.Items.Add(this.btnSelectFirstCellInTables);
            this.groupSelectTables.Label = "Select Tables";
            this.groupSelectTables.Name = "groupSelectTables";
            // 
            // btnSelectAllTables
            // 
            this.btnSelectAllTables.Description = "Select all tables in document";
            this.btnSelectAllTables.Image = global::WordSelectionTool.Properties.Resources.LinkedTableGroup_16x;
            this.btnSelectAllTables.KeyTip = "SAT";
            this.btnSelectAllTables.Label = "Select All Tables";
            this.btnSelectAllTables.Name = "btnSelectAllTables";
            this.btnSelectAllTables.ScreenTip = "Screen Tip";
            this.btnSelectAllTables.ShowImage = true;
            this.btnSelectAllTables.SuperTip = "Super Tip";
            this.btnSelectAllTables.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectAllTables_Click);
            // 
            // btnSelectFirstRow
            // 
            this.btnSelectFirstRow.Image = global::WordSelectionTool.Properties.Resources.Datalist_16x;
            this.btnSelectFirstRow.Label = "Select First Row";
            this.btnSelectFirstRow.Name = "btnSelectFirstRow";
            this.btnSelectFirstRow.ShowImage = true;
            this.btnSelectFirstRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectFirstRow_Click);
            // 
            // btnSelectLastRow
            // 
            this.btnSelectLastRow.Image = global::WordSelectionTool.Properties.Resources.BottomRowOfFourRows_16x;
            this.btnSelectLastRow.Label = "Select Last Row";
            this.btnSelectLastRow.Name = "btnSelectLastRow";
            this.btnSelectLastRow.ShowImage = true;
            this.btnSelectLastRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectLastRow_Click);
            // 
            // btnSelectFirstColumn
            // 
            this.btnSelectFirstColumn.Image = global::WordSelectionTool.Properties.Resources.LeftColumnOfFourColumns_16x;
            this.btnSelectFirstColumn.Label = "Select First Column";
            this.btnSelectFirstColumn.Name = "btnSelectFirstColumn";
            this.btnSelectFirstColumn.ShowImage = true;
            this.btnSelectFirstColumn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectFirstColumn_Click);
            // 
            // btnSelectLastColumn
            // 
            this.btnSelectLastColumn.Image = global::WordSelectionTool.Properties.Resources.RightColumnOfFourColumns_16x;
            this.btnSelectLastColumn.Label = "Select Last Column";
            this.btnSelectLastColumn.Name = "btnSelectLastColumn";
            this.btnSelectLastColumn.ShowImage = true;
            this.btnSelectLastColumn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectLastColumn_Click);
            // 
            // btnSelect1RowTables
            // 
            this.btnSelect1RowTables.Image = global::WordSelectionTool.Properties.Resources.Application_16x;
            this.btnSelect1RowTables.Label = "Select 1 Row Tables";
            this.btnSelect1RowTables.Name = "btnSelect1RowTables";
            this.btnSelect1RowTables.ScreenTip = "Select single row tables";
            this.btnSelect1RowTables.ShowImage = true;
            this.btnSelect1RowTables.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelect1RowTables_Click);
            // 
            // btnSelectMultiRowTables
            // 
            this.btnSelectMultiRowTables.Image = global::WordSelectionTool.Properties.Resources.TopOf3Rows_16x;
            this.btnSelectMultiRowTables.Label = "Select Multi Row Tables";
            this.btnSelectMultiRowTables.Name = "btnSelectMultiRowTables";
            this.btnSelectMultiRowTables.ShowImage = true;
            this.btnSelectMultiRowTables.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectMultiRowTables_Click);
            // 
            // btnSelectFirstCellInTables
            // 
            this.btnSelectFirstCellInTables.Image = global::WordSelectionTool.Properties.Resources.Select1CellInTable8;
            this.btnSelectFirstCellInTables.Label = "Select 1st Cell In Tables";
            this.btnSelectFirstCellInTables.Name = "btnSelectFirstCellInTables";
            this.btnSelectFirstCellInTables.ShowImage = true;
            this.btnSelectFirstCellInTables.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelect1CellInTables_Click);
            // 
            // groupSelectLists
            // 
            this.groupSelectLists.Items.Add(this.btnSelectAllLists);
            this.groupSelectLists.Items.Add(this.btnSelectNumericLists);
            this.groupSelectLists.Items.Add(this.btnSelectBulletLists);
            this.groupSelectLists.Label = "Select Lists";
            this.groupSelectLists.Name = "groupSelectLists";
            // 
            // btnSelectAllLists
            // 
            this.btnSelectAllLists.Image = global::WordSelectionTool.Properties.Resources.ListView_16x;
            this.btnSelectAllLists.Label = "Select All Lists";
            this.btnSelectAllLists.Name = "btnSelectAllLists";
            this.btnSelectAllLists.ShowImage = true;
            this.btnSelectAllLists.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectAllLists_Click);
            // 
            // btnSelectNumericLists
            // 
            this.btnSelectNumericLists.Image = global::WordSelectionTool.Properties.Resources.OrderedList_16x;
            this.btnSelectNumericLists.Label = "Select Numeric Lists";
            this.btnSelectNumericLists.Name = "btnSelectNumericLists";
            this.btnSelectNumericLists.ShowImage = true;
            this.btnSelectNumericLists.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectNumericLists_Click);
            // 
            // btnSelectBulletLists
            // 
            this.btnSelectBulletLists.Image = global::WordSelectionTool.Properties.Resources.BulletList_16x;
            this.btnSelectBulletLists.Label = "Select Bullet Lists";
            this.btnSelectBulletLists.Name = "btnSelectBulletLists";
            this.btnSelectBulletLists.ShowImage = true;
            this.btnSelectBulletLists.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectBulletLists_Click);
            // 
            // groupSelectObjects
            // 
            this.groupSelectObjects.Items.Add(this.btnSelectCurrentPage);
            this.groupSelectObjects.Items.Add(this.btnSelectFormulas);
            this.groupSelectObjects.Items.Add(this.btnSelectShapes);
            this.groupSelectObjects.Label = "Select Objects";
            this.groupSelectObjects.Name = "groupSelectObjects";
            // 
            // btnSelectCurrentPage
            // 
            this.btnSelectCurrentPage.Image = global::WordSelectionTool.Properties.Resources.Document_16x;
            this.btnSelectCurrentPage.Label = "Select Current Page";
            this.btnSelectCurrentPage.Name = "btnSelectCurrentPage";
            this.btnSelectCurrentPage.ShowImage = true;
            this.btnSelectCurrentPage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectCurrentPage_Click);
            // 
            // btnSelectFormulas
            // 
            this.btnSelectFormulas.Image = global::WordSelectionTool.Properties.Resources.AutoSum_16x;
            this.btnSelectFormulas.Label = "Select Formulas";
            this.btnSelectFormulas.Name = "btnSelectFormulas";
            this.btnSelectFormulas.ShowImage = true;
            this.btnSelectFormulas.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectFormulas_Click);
            // 
            // btnSelectShapes
            // 
            this.btnSelectShapes.Enabled = false;
            this.btnSelectShapes.Image = global::WordSelectionTool.Properties.Resources.AbstractCube_16x;
            this.btnSelectShapes.Label = "Select Shapes";
            this.btnSelectShapes.Name = "btnSelectShapes";
            this.btnSelectShapes.ShowImage = true;
            this.btnSelectShapes.Visible = false;
            this.btnSelectShapes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectPictures_Click);
            // 
            // groupHelp
            // 
            this.groupHelp.Items.Add(this.btnAbout);
            this.groupHelp.Label = "Help";
            this.groupHelp.Name = "groupHelp";
            // 
            // btnAbout
            // 
            this.btnAbout.Image = global::WordSelectionTool.Properties.Resources.InfoRule_16x;
            this.btnAbout.Label = "About";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.ShowImage = true;
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabSelectTools);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabSelectTools.ResumeLayout(false);
            this.tabSelectTools.PerformLayout();
            this.groupSelectTables.ResumeLayout(false);
            this.groupSelectTables.PerformLayout();
            this.groupSelectLists.ResumeLayout(false);
            this.groupSelectLists.PerformLayout();
            this.groupSelectObjects.ResumeLayout(false);
            this.groupSelectObjects.PerformLayout();
            this.groupHelp.ResumeLayout(false);
            this.groupHelp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabSelectTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSelectTables;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectAllTables;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectFirstRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectLastRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectFirstColumn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectLastColumn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelect1RowTables;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectCurrentPage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectShapes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectAllLists;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectFormulas;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSelectObjects;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectMultiRowTables;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectFirstCellInTables;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSelectLists;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectNumericLists;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectBulletLists;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
