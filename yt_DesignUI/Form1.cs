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

namespace yt_DesignUI
{
    public partial class Form1 : ShadowedForm
    {
        public Form1()
        {
            InitializeComponent();


            //buttonAnim.Value = button1.Width;
            //Animator.Start();


            if (cmbStyle.Items.Count == 0)
            {
                EgoldsFormStyle.fStyle selectedStyle = egoldsFormStyle1.FormStyle;
                cmbStyle.DataSource = Enum.GetValues(typeof(EgoldsFormStyle.fStyle));
                cmbStyle.SelectedItem = selectedStyle;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void cmbStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            egoldsFormStyle1.FormStyle = (EgoldsFormStyle.fStyle)cmbStyle.SelectedItem;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void yt_Button3_Click(object sender, EventArgs e)
        {
            Hide();
            yt_DesignUI.Form3 f3 = new yt_DesignUI.Form3();
            f3.ShowDialog();
            Close();
        }

        private void yt_Button6_Click(object sender, EventArgs e)
        {
            Hide();
            yt_DesignUI.Form2 f2 = new yt_DesignUI.Form2();
            f2.ShowDialog();
            Close();
        }
    }
}
