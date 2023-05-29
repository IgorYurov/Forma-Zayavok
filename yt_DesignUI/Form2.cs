using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using yt_DesignUI.Components;
using yt_DesignUI.Controls;
using System.Data.OleDb;
using Word = Microsoft.Office.Interop.Word;

namespace yt_DesignUI
{
    public partial class Form2 : Form
    {
        public static string connectString = "Provider=Microsoft.ACE.OLEDB.12.0;" + @"Data Source=|DataDirectory|\\BazaJiEst.accdb";
        static OleDbConnection myConnection = new OleDbConnection(connectString);
        OleDbDataAdapter DataAdapter = new OleDbDataAdapter("SELECT * FROM Tablica12", myConnection);
        DataSet dt = new DataSet();

        public Form2()
        {
            InitializeComponent();
            DataAdapter.Fill(dt);

            //panel1.MouseWheel += OnMouseWheel;
            //this.MouseWheel += new MouseEventHandler(panel1_MouseWheel);
            //this.panel1.MouseWheel += System.Windows.Forms.MouseEventHandler(this.panel1_MouseWheel);

            //Animator.Start();


            if (cmbStyle.Items.Count == 0)
            {
                EgoldsFormStyle.fStyle selectedStyle = egoldsFormStyle1.FormStyle;
                cmbStyle.DataSource = Enum.GetValues(typeof(EgoldsFormStyle.fStyle));
                cmbStyle.SelectedItem = selectedStyle;
            }
        }

        private void panel1_MouseEnter(object sender, EventArgs e)
        {
            panel1.Focus();
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "bazaJiEstDataSet2.Tablica12". При необходимости она может быть перемещена или удалена.
            this.tablica12TableAdapter2.Fill(this.bazaJiEstDataSet2.Tablica12);

            dataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView3.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView3.DataSource = dt.Tables[0].DefaultView;

        }

        //public void OnMouseWheel(object sender, MouseEvents e)
        //{
        //    if (mouseOverPanel)
        //    {
        //        if (e.Delta < 0)
        //            vScrollBar1.Value++;
        //        else
        //            vScrollBar1.Value--;
        //    }
        //}

        private void cmbStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            egoldsFormStyle1.FormStyle = (EgoldsFormStyle.fStyle)cmbStyle.SelectedItem;
        }

        //private void dataGridView3_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        //{
        //}

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void egoldsGoogleTextBox5_Click(object sender, EventArgs e)
        {

        }

        private void yt_Button3_Click(object sender, EventArgs e)
        {
            Hide();
            yt_DesignUI.Form3 f3 = new yt_DesignUI.Form3();
            f3.ShowDialog();
            Close();
        }

        private void egoldsGoogleTextBox1_Click(object sender, EventArgs e)
        {
              
        }

        private void egoldsGoogleTextBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            //if (comboBox1.Text == "Заявителю")
            //{
            //    BindingSource bs = new BindingSource();
            //    bs.DataSource = dataGridView3.DataSource;
            //    bs.Filter = string.Format("CONVERT(" + dataGridView3.Columns[1].DataPropertyName + ", System.String) like '%" + textBox1.Text.Replace("'", "''") + "%'");
            //    dataGridView3.DataSource = bs;
            //}
            //if (comboBox1.Text == "Регистрационному номеру")
            //{
            //    BindingSource bs = new BindingSource();
            //    bs.DataSource = dataGridView3.DataSource;
            //    bs.Filter = string.Format("CONVERT(" + dataGridView3.Columns[2].DataPropertyName + ", System.String) like '%" + textBox1.Text.Replace("'", "''") + "%'");
            //    dataGridView3.DataSource = bs;
            //}
        }

        //private void egoldsGoogleTextBox1_Click_1(object sender, EventArgs e)
        //{
        //}

        private void dataGridView3_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            Hide();
            //Form2 f2 = new Form2();
            yt_DesignUI.Form3 f3 = new yt_DesignUI.Form3();
            f3.textBox1.Text = dataGridView3.CurrentRow.Cells[1].Value.ToString();
            f3.egoldsGoogleTextBox1.Text = dataGridView3.CurrentRow.Cells[2].Value.ToString();
            f3.egoldsGoogleTextBox2.Text = dataGridView3.CurrentRow.Cells[3].Value.ToString();
            f3.egoldsGoogleTextBox4.Text = dataGridView3.CurrentRow.Cells[4].Value.ToString();
            f3.egoldsGoogleTextBox5.Text = dataGridView3.CurrentRow.Cells[5].Value.ToString();
            f3.egoldsGoogleTextBox3.Text = dataGridView3.CurrentRow.Cells[6].Value.ToString();
            f3.comboBox6.Text = dataGridView3.CurrentRow.Cells[7].Value.ToString();
            f3.comboBox3.Text = dataGridView3.CurrentRow.Cells[8].Value.ToString();
            f3.egoldsGoogleTextBox6.Text = dataGridView3.CurrentRow.Cells[9].Value.ToString();
            f3.egoldsGoogleTextBox7.Text = dataGridView3.CurrentRow.Cells[10].Value.ToString();
            f3.egoldsGoogleTextBox8.Text = dataGridView3.CurrentRow.Cells[11].Value.ToString();
            f3.egoldsGoogleTextBox9.Text = dataGridView3.CurrentRow.Cells[12].Value.ToString();
            f3.egoldsGoogleTextBox10.Text = dataGridView3.CurrentRow.Cells[13].Value.ToString();
            f3.comboBox4.Text = dataGridView3.CurrentRow.Cells[14].Value.ToString();
            f3.comboBox5.Text = dataGridView3.CurrentRow.Cells[15].Value.ToString();
            f3.egoldsGoogleTextBox11.Text = dataGridView3.CurrentRow.Cells[16].Value.ToString();
            f3.egoldsGoogleTextBox12.Text = dataGridView3.CurrentRow.Cells[17].Value.ToString();
            //textBox11.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
            //textBox12.Text = dataGridView1.CurrentRow.Cells[19].Value.ToString();
            //textBox13.Text = dataGridView1.CurrentRow.Cells[20].Value.ToString();
            //textBox14.Text = dataGridView1.CurrentRow.Cells[21].Value.ToString();
            //f2.comboBox7.Text = dataGridView3.CurrentRow.Cells[22].Value.ToString();
            f3.comboBox11.Text = dataGridView3.CurrentRow.Cells[23].Value.ToString();
            f3.ShowDialog();
            Close();
        }

        private void egoldsGoogleTextBox18_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Заявителю")
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView3.DataSource;
                bs.Filter = string.Format("CONVERT(" + dataGridView3.Columns[1].DataPropertyName + ", System.String) like '%" + egoldsGoogleTextBox18.Text.Replace("'", "''") + "%'");
                dataGridView3.DataSource = bs;
            }
            if (comboBox1.Text == "Регистрационному номеру")
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView3.DataSource;
                bs.Filter = string.Format("CONVERT(" + dataGridView3.Columns[2].DataPropertyName + ", System.String) like '%" + egoldsGoogleTextBox18.Text.Replace("'", "''") + "%'");
                dataGridView3.DataSource = bs;
            }
        }

        //private void egoldsGoogleTextBox2_Click(object sender, EventArgs e)
        //{
        //}
    }
}
