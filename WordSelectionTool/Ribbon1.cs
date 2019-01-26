using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace WordSelectionTool
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnSelectAllTables_Click(object sender, RibbonControlEventArgs e)
        {
            //System.Windows.Forms.MessageBox.Show("Button Clicked!");
            if (!Globals.ThisAddIn.selectAllTables())
                System.Windows.Forms.MessageBox.Show(DOC_DOES_NOT_CONTAIN_TABLES);
        }

        /*private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Before create table!");

            Globals.ThisAddIn.addTable();

            System.Windows.Forms.MessageBox.Show("After create table!");

        }*/

        private void btnSelectFirstRow_Click(object sender, RibbonControlEventArgs e)
        {
            if(!Globals.ThisAddIn.selectFirstRowInTables())
                System.Windows.Forms.MessageBox.Show(DOC_DOES_NOT_CONTAIN_TABLES);
        }

        private void btnSelectLastRow_Click(object sender, RibbonControlEventArgs e)
        {
            if(!Globals.ThisAddIn.selectLastRowInTables())
                System.Windows.Forms.MessageBox.Show(DOC_DOES_NOT_CONTAIN_TABLES);
        }

        private void btnSelectFirstColumn_Click(object sender, RibbonControlEventArgs e)
        {
            if(!Globals.ThisAddIn.selectFirstColumnInTables())
                System.Windows.Forms.MessageBox.Show(DOC_DOES_NOT_CONTAIN_TABLES);
        }

        private void btnSelectLastColumn_Click(object sender, RibbonControlEventArgs e)
        {
            if(!Globals.ThisAddIn.selectLastColumnInTables())
                System.Windows.Forms.MessageBox.Show(DOC_DOES_NOT_CONTAIN_TABLES);
        }

        private void btnSelect1RowTables_Click(object sender, RibbonControlEventArgs e)
        {
            if(!Globals.ThisAddIn.select1RowTables())
                System.Windows.Forms.MessageBox.Show("This document does not contain single row tables");
        }

        private void btnSelectCurrentPage_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.selectCurrentPage();
        }

        private void btnSelectPictures_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.selectShapes();
        }

        private void btnSelectAllLists_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.selectAllLists())
                System.Windows.Forms.MessageBox.Show("This document does not contain any lists");
        }

        private void btnSelectFormulas_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.selectFormulas())
                System.Windows.Forms.MessageBox.Show("This document does not contain any formulas");
        }

        private void btnSelectMultiRowTables_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.selectMultiRowTables())
                System.Windows.Forms.MessageBox.Show("This document does not contain multi row tables");
        }

        private void btnSelect1CellInTables_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.selectFirstCellInTables())
                System.Windows.Forms.MessageBox.Show(DOC_DOES_NOT_CONTAIN_TABLES);
        }

        private void btnSelectNumericLists_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.selectNumericLists())
                System.Windows.Forms.MessageBox.Show("This document does not contain any numeric lists"); 
        }

        private void btnSelectBulletLists_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.selectBulletLists())
                System.Windows.Forms.MessageBox.Show("This document does not contain any bullet lists");
        }

        /*private void btnSelectBullets_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }*/

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            using(AboutBox ab = new AboutBox())
            {
                ab.ShowDialog();
            }
        }

        private void btnRegister_Click(object sender, RibbonControlEventArgs e)
        {
            ;
        }

        private const string DOC_DOES_NOT_CONTAIN_TABLES = "This document does not contain tables";
       
    }
}
