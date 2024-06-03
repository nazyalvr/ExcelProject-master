using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;

namespace excel
{
    public partial class Form1 : Form
    {
        bool expressionview;
        string actualpath = "";

        public Form1(string[] args)
        {
            InitializeComponent();
            CellCount.Instance.Providetable(table);
            SetupTableSize(10, 10);
            expressionview = false;
            if (args.Length == 1)
            {
                LoadSavedTable(args[0]);
            }
        }

        private void SetupTableSize(int colnum, int rownum)
        {
            typeof(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic |
                BindingFlags.Instance | BindingFlags.SetProperty, null, table, new object[] { true });
            table.AllowUserToAddRows = false;
            table.ColumnCount = colnum;
            table.RowCount = rownum;
            UpdateTable();
        }

        private void Savetable(string path)
        {
            actualpath = path;
            table.EndEdit();
            DataTable datatab = new DataTable("base");
            foreach (DataGridViewColumn col in table.Columns)
            {
                datatab.Columns.Add(col.Index.ToString());
            }
            foreach(DataGridViewRow ro in table.Rows)
            {
                DataRow dtnewrow = datatab.NewRow();
                foreach(DataColumn col in datatab.Columns)
                {
                    dtnewrow[col.ColumnName] = CellCount.Instance.TakeCell(ro.Cells[Parser.parse(col.ColumnName)]).expression;
                }
            }
        }
        private void LoadSavedTable(string path)
        {
            actualpath = path;
            DataSet dataset = new DataSet();
            dataset.ReadXml(path);
            DataTable datatab = dataset.Tables[0];
            table.ColumnCount = datatab.Columns.Count;
            table.RowCount = datatab.Rows.Count;
            foreach(DataGridViewRow ro in table.Rows)
            {
                foreach(DataGridViewCell tablecell in ro.Cells)
                {
                    tablecell.Tag = new Cell(tablecell, datatab.Rows[tablecell.RowIndex][tablecell.ColumnIndex].ToString()); 
                }
            }
            UpdateTable();
            UpdateCells();
        }

        private void ChangeView()
        {
            foreach(DataGridViewRow ro in table.Rows)
            {
                foreach(DataGridViewCell tablecell in ro.Cells)
                {
                    Cell cell = CellCount.Instance.TakeCell(tablecell);
                    if (!expressionview)
                    {
                        tablecell.Value = cell.Expression;
                    }
                    else
                    {
                        if (cell.Expression == "")
                        {
                            tablecell.Value = cell.Expression;
                        }
                        else
                        {
                            tablecell.Value = cell.Value;
                        }
                    }
                }
            }
            expressionview = !expressionview;
        }


        private void UpdateTable()
        {
            foreach(DataGridViewColumn col in table.Columns)
            {
                col.HeaderText = "B" + (col.Index + 1);
                col.MinimumWidth = 80;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            foreach(DataGridViewRow row in table.Rows)
            {
                row.HeaderCell.Value = "A" + (row.Index + 1);
            }

            foreach (DataGridViewRow row in table.Rows)
            {
                foreach(DataGridViewCell tablecell in row.Cells)
                {
                    if (tablecell == null)
                    {
                        tablecell.Tag = new Cell(tablecell, "");
                    }
                }    
            }
        }

        private void Addrow()
        {
            table.RowCount++;
            UpdateTable();
        }

        private void DeleteRow()
        {
            DialogResult res = MessageBox.Show("Do you want to  kill this row?(((", "Shoot it down", MessageBoxButtons.YesNo);
            if (res == DialogResult.Yes)
            {
                if (table.RowCount >= 0)
                {
                    DataGridViewCell tablecell = table.SelectedCells[0];
                    table.Rows.RemoveAt(tablecell.RowIndex);
                }
                UpdateTable();
                UpdateCells();
            }
            else { }
        }

        public void AddColumn()
        {
            table.ColumnCount++;
            UpdateTable();
        }

        private void DeleteColumn()
        {
            DialogResult res = MessageBox.Show("Do you want to  kill this column?(((", "Shoot it down", MessageBoxButtons.YesNo);
            if (res == DialogResult.Yes)
            {
                if (table.ColumnCount >= 0)
                {
                    DataGridViewCell tablecell = table.SelectedCells[0];
                    table.Columns.RemoveAt(tablecell.ColumnIndex);
                }
                UpdateTable();
                UpdateCells();
            }
            else { }
        }


        public void UpdateCells()
        {
            foreach(DataGridViewRow row in table.Rows)
            {
                foreach (DataGridViewCell tablecell in row.Cells)
                {
                    Cell cell = (Cell)tablecell.Tag;
                    cell.process();    
                    if (!expressionview)
                    {
                        if (cell.Expression == "")
                        {
                            tablecell.Value = cell.Expression;
                        }
                        else 
                            tablecell.Value = cell.Value.ToString();
                    }
                }
            }
        }
        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void menuToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Програма розроблена Лаврентюком Назаром. Підтримує операції:'+','-','*','/','^', унарні операції, mmax(n values), mmin(n values), not, логічні операції тощо.", "Info");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        { 
            DialogResult result = MessageBox.Show("Do you want to save table?", "Save table", MessageBoxButtons.YesNoCancel);
            if (result == DialogResult.Yes)
            {
                if (actualpath != "")
                {
                    Savetable(actualpath);
                }
                else
                {
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        Savetable(saveFileDialog1.FileName);
                    };
                }
            }
            else if (result == DialogResult.Cancel)
            {
                e.Cancel = true;
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            DeleteRow();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            AddColumn();
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            DeleteColumn();
        }

        private void toolStripButton0_Click(object sender, EventArgs e)
        {
            ChangeView();
        }

        public void tableCellendendit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            Cell cell = CellCount.Instance.TakeCell(e.RowIndex, e.ColumnIndex);
            DataGridViewCell tablecell = cell.Example;
            string firstexpression = cell.Expression;
            if (tablecell.Value != null)
            {
                cell.Expression = tablecell.Value.ToString();
                try 
                {
                    UpdateCells();
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.GetType().ToString());
                    cell.Expression = firstexpression;
                    UpdateCells();
                }
            }
            else
            {
                cell.Expression = "";
            }
        }


        private void tablecelldoubleclick(object sender, DataGridViewCellEventArgs e)
        {
            table.BeginEdit(true);
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                LoadSavedTable(openFileDialog1.FileName);
            };
        }

        private void saveToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (actualpath != "")
            {
                Savetable(actualpath);
            }
            else
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Savetable(saveFileDialog1.FileName); 
                };
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (this.Width > SystemInformation.VirtualScreen.Width)
            {
                this.Width = SystemInformation.VirtualScreen.Width;
                this.Left = 0;
            }
            if (this.Height > (SystemInformation.VirtualScreen.Height * 9) / 10)
            {
                this.Height = (SystemInformation.VirtualScreen.Height * 9) / 10;
                this.Top = 0;
            }
        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you want to save table?", "Save table", MessageBoxButtons.YesNoCancel);
            if (result == DialogResult.Yes)
            {
                if (actualpath != "")
                {
                    Savetable(actualpath);
                }
                else
                {
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        Savetable(saveFileDialog1.FileName);
                    };
                }
            };
        }

        private void tableCellbeginedit(object sender, DataGridViewCellCancelEventArgs e)
        {
            Cell cell = CellCount.Instance.TakeCell(e.RowIndex, e.ColumnIndex);
            DataGridViewCell tablecell = cell.Example;
            tablecell.Value = cell.Expression;
        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            Addrow();
        }
    }
}
