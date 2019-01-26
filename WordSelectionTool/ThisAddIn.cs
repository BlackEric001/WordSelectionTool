using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

using System.IO;

namespace WordSelectionTool
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public bool selectAllTables()
        {
            ////Курсор должен быть за пределами таблицы. Иначе падает.
            if (Application.ActiveDocument.Tables.Count > 0)
            {
                Application.ActiveDocument.Range(0, 0).Select();

                //Application.ScreenUpdating = false; //Если включить, то не работает в 2016 ворде

                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                foreach (Word.Table table in Application.ActiveDocument.Tables)
                    table.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                Application.ActiveDocument.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                //Application.ScreenUpdating = true; //Если включить, то не работает в 2016 ворде

                return true;
            }
            else
                return false;
        }

        public bool selectFirstRowInTables()
        {
            if (Application.ActiveDocument.Tables.Count > 0)
            {
                Application.ActiveDocument.Range(0, 0).Select();

                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                foreach (Word.Table table in Application.ActiveDocument.Tables)
                    table.Rows[1].Range.Editors.Add(Word.WdEditorType.wdEditorEveryone); //first row

                Application.ActiveDocument.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                return true;
            }
            else
                return false;
        }

        public bool selectLastRowInTables()
        {
            if (Application.ActiveDocument.Tables.Count > 0)
            {
                Application.ActiveDocument.Range(0, 0).Select();

                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                foreach (Word.Table table in Application.ActiveDocument.Tables)
                    table.Rows[table.Rows.Count].Range.Editors.Add(Word.WdEditorType.wdEditorEveryone); //last row

                Application.ActiveDocument.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                return true;
            }
            else
                return false;
        }

        public bool selectFirstColumnInTables()
        {
            if (Application.ActiveDocument.Tables.Count > 0)
            {
                Application.ActiveDocument.Range(0, 0).Select();

                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                foreach (Word.Table table in Application.ActiveDocument.Tables)
                    foreach (Word.Cell cell in table.Columns.First.Cells)  //First column
                        cell.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                Application.ActiveDocument.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                return true;
            }
            else
                return false;
        }

        public bool selectLastColumnInTables()
        {
            if (Application.ActiveDocument.Tables.Count > 0)
            {
                Application.ActiveDocument.Range(0, 0).Select();

                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                foreach (Word.Table table in Application.ActiveDocument.Tables)
                    foreach (Word.Cell cell in table.Columns.Last.Cells)  //Last column
                        cell.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                Application.ActiveDocument.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                return true;
            }
            else
                return false;
        }

        public bool selectFirstCellInTables()
        {
            if (Application.ActiveDocument.Tables.Count > 0)
            {
                Application.ActiveDocument.Range(0, 0).Select();

                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                foreach (Word.Table table in Application.ActiveDocument.Tables)
                    table.Columns.First.Cells[1].Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);  //first cell in table

                Application.ActiveDocument.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                return true;
            }
            else
                return false;
        }

        public bool select1RowTables()
        {
            if (Application.ActiveDocument.Tables.Count > 0)
            {
                Application.ActiveDocument.Range(0, 0).Select();

                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                foreach (Word.Table table in Application.ActiveDocument.Tables)
                    if (table.Rows.Count == 1) // single line tables
                        table.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                Application.ActiveDocument.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                return true;
            }
            else
                return false;
        }

        public bool selectMultiRowTables()
        {
            if (Application.ActiveDocument.Tables.Count > 0)
            {
                Application.ActiveDocument.Range(0, 0).Select();

                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                foreach (Word.Table table in Application.ActiveDocument.Tables)
                    if (table.Rows.Count > 1) //multiline tables
                        table.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                Application.ActiveDocument.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                return true;
            }
            else
                return false;
        }

        public bool selectCurrentPage()
        {
            Application.ActiveDocument.Bookmarks[@"\page"].Range.Select();
            return true;
        }

        public bool selectShapes()
        {
            if (Application.ActiveDocument.InlineShapes.Count > 0)
                Application.ActiveDocument.Range(0, 0).Select();
            //Application.Selection.Document.Content.Select(); //select all
            //Application.Selection.Document.Shapes.SelectAll(); //select all shapes
            //Application.Selection.Document.Lists[1].Range.Select();

            Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

            foreach (Word.InlineShape ishape in Application.ActiveDocument.InlineShapes)
                if (ishape.Type == Microsoft.Office.Interop.Word.WdInlineShapeType.wdInlineShapePicture)
                    ishape.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

            Application.ActiveDocument.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
            Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

            return true;
        }

        public bool selectAllLists()
        {
            if (Application.ActiveDocument.Lists.Count > 0)
            {
                Application.ActiveDocument.Range(0, 0).Select();
                //Application.Selection.Document.Content.Select(); //select all
                //Application.Selection.Document.Lists[1].Range.Select();

                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                foreach (Word.List list in Application.ActiveDocument.Lists)
                {
                    log(list.Range.ListFormat.ListType.ToString());
                    list.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);
                }

                Application.ActiveDocument.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                return true;
            }
            else
                return false;
        }

        public bool selectBulletLists()
        {
            bool res = false;
            if (Application.ActiveDocument.Lists.Count > 0)
            {
                foreach (Word.List list in Application.ActiveDocument.Lists)
                    if (list.Range.ListFormat.ListType == Word.WdListType.wdListBullet)
                    {
                        Application.ActiveDocument.Range(0, 0).Select();
                        break;
                    }

                //Application.Selection.Document.Content.Select(); //select all
                //Application.Selection.Document.Lists[1].Range.Select();

                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                foreach (Word.List list in Application.ActiveDocument.Lists)
                {
                    log(list.Range.ListFormat.ListType.ToString());
                    //log(list.Range.);
                    if (list.Range.ListFormat.ListType == Word.WdListType.wdListBullet)
                    {
                        list.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);
                        res = true;
                    }
                }

                Application.ActiveDocument.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
            }
            return res;
        }

        public bool selectNumericLists()
        {
            bool res = false;
            if (Application.ActiveDocument.Lists.Count > 0)
            {
                foreach (Word.List list in Application.ActiveDocument.Lists)
                    //if (list.Range.ListFormat.ListType == Word.WdListType.wdListSimpleNumbering)
                    if (
                        list.Range.ListFormat.ListType == Word.WdListType.wdListSimpleNumbering ||
                        //list.Range.ListFormat.ListType == Word.WdListType.wdListMixedNumbering || 
                        list.Range.ListFormat.ListType == Word.WdListType.wdListOutlineNumbering
                        )
                    {
                        Application.ActiveDocument.Range(0, 0).Select();
                        break;
                    }

                //Application.Selection.Document.Content.Select(); //select all
                //Application.Selection.Document.Lists[1].Range.Select();

                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                foreach (Word.List list in Application.ActiveDocument.Lists)
                {
                    log(list.Range.ListFormat.ListType.ToString());
                    if (
                        list.Range.ListFormat.ListType == Word.WdListType.wdListSimpleNumbering ||
                    //list.Range.ListFormat.ListType == Word.WdListType.wdListMixedNumbering ||
                    list.Range.ListFormat.ListType == Word.WdListType.wdListOutlineNumbering
                        )
                    {
                        list.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);
                        res = true;
                    }
                }

                Application.ActiveDocument.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);


            }
            return res;
        }

        public bool selectFormulas()
        {
            if (Application.ActiveDocument.OMaths.Count > 0)
            {
                Application.ActiveDocument.Range(0, 0).Select();

                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                foreach (Word.OMath math in Application.ActiveDocument.OMaths)
                    math.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                Application.ActiveDocument.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone);
                Application.ActiveDocument.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

                return true;
            }
            else
                return false;
        }

        private void log(string data)
        {
            #if DEBUG
                using (StreamWriter outputFile = new StreamWriter(Environment.CurrentDirectory + @"\log.txt", true))
                {
                    outputFile.WriteLine(data);
                }
            #endif
        }

        public void addTable()
        {
            Word.Range tableLocation =
                this.Application.ActiveDocument.Range(0, 0);
            this.Application.ActiveDocument.Tables.Add(
                tableLocation, 5, 5);

            this.Application.ActiveDocument.Tables[1].Range.Font.Size = 8;
            this.Application.ActiveDocument.Tables[1].Range.Cells.Borders.InsideColor = Word.WdColor.wdColorBlack;
            //this.Application.ActiveDocument.Tables[1].set_Style("Table Grid 8");
            //How to get list of table styles?
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
