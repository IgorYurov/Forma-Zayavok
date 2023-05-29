
namespace yt_DesignUI
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.egoldsFormStyle1 = new yt_DesignUI.Components.EgoldsFormStyle(this.components);
            this.cmbStyle = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.yt_Button6 = new yt_DesignUI.yt_Button();
            this.yt_Button3 = new yt_DesignUI.yt_Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // egoldsFormStyle1
            // 
            this.egoldsFormStyle1.AllowUserResize = true;
            this.egoldsFormStyle1.BackColor = System.Drawing.SystemColors.HighlightText;
            this.egoldsFormStyle1.ContextMenuForm = null;
            this.egoldsFormStyle1.ControlBoxButtonsWidth = 60;
            this.egoldsFormStyle1.EnableControlBoxIconsLight = true;
            this.egoldsFormStyle1.EnableControlBoxMouseLight = true;
            this.egoldsFormStyle1.Form = this;
            this.egoldsFormStyle1.FormStyle = yt_DesignUI.Components.EgoldsFormStyle.fStyle.SimpleDark;
            this.egoldsFormStyle1.HeaderColor = System.Drawing.Color.Violet;
            this.egoldsFormStyle1.HeaderColorAdditional = System.Drawing.Color.RoyalBlue;
            this.egoldsFormStyle1.HeaderColorGradientEnable = true;
            this.egoldsFormStyle1.HeaderColorGradientMode = System.Drawing.Drawing2D.LinearGradientMode.Horizontal;
            this.egoldsFormStyle1.HeaderHeight = 38;
            this.egoldsFormStyle1.HeaderImage = null;
            this.egoldsFormStyle1.HeaderTextColor = System.Drawing.Color.White;
            this.egoldsFormStyle1.HeaderTextFont = new System.Drawing.Font("Segoe UI", 9.75F);
            // 
            // cmbStyle
            // 
            this.cmbStyle.FormattingEnabled = true;
            this.cmbStyle.Location = new System.Drawing.Point(1734, 13);
            this.cmbStyle.Name = "cmbStyle";
            this.cmbStyle.Size = new System.Drawing.Size(121, 25);
            this.cmbStyle.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel1.Controls.Add(this.yt_Button6);
            this.panel1.Controls.Add(this.yt_Button3);
            this.panel1.Location = new System.Drawing.Point(0, -1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(789, 377);
            this.panel1.TabIndex = 1;
            // 
            // yt_Button6
            // 
            this.yt_Button6.BackColor = System.Drawing.Color.Green;
            this.yt_Button6.BackColorAdditional = System.Drawing.Color.Lime;
            this.yt_Button6.BackColorGradientEnabled = true;
            this.yt_Button6.BackColorGradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
            this.yt_Button6.BorderColor = System.Drawing.Color.DimGray;
            this.yt_Button6.BorderColorEnabled = true;
            this.yt_Button6.BorderColorOnHover = System.Drawing.Color.Tomato;
            this.yt_Button6.BorderColorOnHoverEnabled = false;
            this.yt_Button6.Cursor = System.Windows.Forms.Cursors.Hand;
            this.yt_Button6.Font = new System.Drawing.Font("Verdana", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.yt_Button6.ForeColor = System.Drawing.Color.White;
            this.yt_Button6.Location = new System.Drawing.Point(213, 206);
            this.yt_Button6.Name = "yt_Button6";
            this.yt_Button6.RippleColor = System.Drawing.Color.Black;
            this.yt_Button6.Rounding = 60;
            this.yt_Button6.RoundingEnable = true;
            this.yt_Button6.Size = new System.Drawing.Size(362, 55);
            this.yt_Button6.TabIndex = 97;
            this.yt_Button6.Text = "Открыть базу";
            this.yt_Button6.TextHover = null;
            this.yt_Button6.UseDownPressEffectOnClick = true;
            this.yt_Button6.UseRippleEffect = true;
            this.yt_Button6.UseZoomEffectOnHover = true;
            this.yt_Button6.Click += new System.EventHandler(this.yt_Button6_Click);
            // 
            // yt_Button3
            // 
            this.yt_Button3.BackColor = System.Drawing.Color.Blue;
            this.yt_Button3.BackColorAdditional = System.Drawing.Color.RoyalBlue;
            this.yt_Button3.BackColorGradientEnabled = true;
            this.yt_Button3.BackColorGradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
            this.yt_Button3.BorderColor = System.Drawing.Color.DimGray;
            this.yt_Button3.BorderColorEnabled = true;
            this.yt_Button3.BorderColorOnHover = System.Drawing.Color.Tomato;
            this.yt_Button3.BorderColorOnHoverEnabled = false;
            this.yt_Button3.Cursor = System.Windows.Forms.Cursors.Hand;
            this.yt_Button3.Font = new System.Drawing.Font("Verdana", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.yt_Button3.ForeColor = System.Drawing.Color.White;
            this.yt_Button3.Location = new System.Drawing.Point(158, 103);
            this.yt_Button3.Name = "yt_Button3";
            this.yt_Button3.RippleColor = System.Drawing.Color.Black;
            this.yt_Button3.Rounding = 60;
            this.yt_Button3.RoundingEnable = true;
            this.yt_Button3.Size = new System.Drawing.Size(471, 55);
            this.yt_Button3.TabIndex = 80;
            this.yt_Button3.Text = "Сформировать новое заявление";
            this.yt_Button3.TextHover = null;
            this.yt_Button3.UseDownPressEffectOnClick = true;
            this.yt_Button3.UseRippleEffect = true;
            this.yt_Button3.UseZoomEffectOnHover = true;
            this.yt_Button3.Click += new System.EventHandler(this.yt_Button3_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(789, 377);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.cmbStyle);
            this.Font = new System.Drawing.Font("Segoe UI", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ForeColor = System.Drawing.SystemColors.Control;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Inter-Sert";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private Components.EgoldsFormStyle egoldsFormStyle1;
        private System.Windows.Forms.ComboBox cmbStyle;
        private System.Windows.Forms.Panel panel1;
        private yt_Button yt_Button3;
        private yt_Button yt_Button6;
    }
}