using System.Windows.Forms;

namespace CloseActivities
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.BtnGetSQL = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.BtnLoadSQL = new System.Windows.Forms.Button();
            this.panel5 = new System.Windows.Forms.Panel();
            this.Button5 = new System.Windows.Forms.Button();
            this.Button4 = new System.Windows.Forms.Button();
            this.Button3 = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.label14 = new System.Windows.Forms.Label();
            this.BtnOpenSapGui = new System.Windows.Forms.Button();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUser = new System.Windows.Forms.TextBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.BtnOpenSAP = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.panel6 = new System.Windows.Forms.Panel();
            this.RadioButton2 = new System.Windows.Forms.RadioButton();
            this.RadioButton1 = new System.Windows.Forms.RadioButton();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel5.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel6.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.BtnGetSQL);
            this.panel1.Location = new System.Drawing.Point(103, 55);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(250, 39);
            this.panel1.TabIndex = 59;
            // 
            // BtnGetSQL
            // 
            this.BtnGetSQL.Location = new System.Drawing.Point(15, 3);
            this.BtnGetSQL.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BtnGetSQL.Name = "BtnGetSQL";
            this.BtnGetSQL.Size = new System.Drawing.Size(220, 34);
            this.BtnGetSQL.TabIndex = 58;
            this.BtnGetSQL.Text = "Choisir le Fichier Excel";
            this.BtnGetSQL.UseVisualStyleBackColor = true;
            this.BtnGetSQL.Click += new System.EventHandler(this.BtnGetSQL_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.BtnLoadSQL);
            this.panel2.Location = new System.Drawing.Point(371, 55);
            this.panel2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(250, 39);
            this.panel2.TabIndex = 60;
            // 
            // BtnLoadSQL
            // 
            this.BtnLoadSQL.Location = new System.Drawing.Point(15, 3);
            this.BtnLoadSQL.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BtnLoadSQL.Name = "BtnLoadSQL";
            this.BtnLoadSQL.Size = new System.Drawing.Size(220, 34);
            this.BtnLoadSQL.TabIndex = 59;
            this.BtnLoadSQL.Text = "Télécharger dans SQL Server";
            this.BtnLoadSQL.UseVisualStyleBackColor = true;
            this.BtnLoadSQL.Click += new System.EventHandler(this.BtnLoadSQL_Click);
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.Button5);
            this.panel5.Controls.Add(this.Button4);
            this.panel5.Controls.Add(this.Button3);
            this.panel5.Controls.Add(this.label7);
            this.panel5.Controls.Add(this.label6);
            this.panel5.Controls.Add(this.label5);
            this.panel5.Controls.Add(this.label4);
            this.panel5.Controls.Add(this.label3);
            this.panel5.Controls.Add(this.label2);
            this.panel5.Controls.Add(this.label1);
            this.panel5.Controls.Add(this.dateTimePicker2);
            this.panel5.Controls.Add(this.dateTimePicker1);
            this.panel5.Controls.Add(this.textBox5);
            this.panel5.Controls.Add(this.textBox4);
            this.panel5.Controls.Add(this.textBox3);
            this.panel5.Controls.Add(this.textBox2);
            this.panel5.Controls.Add(this.textBox1);
            this.panel5.Location = new System.Drawing.Point(103, 171);
            this.panel5.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(788, 420);
            this.panel5.TabIndex = 61;
            // 
            // Button5
            // 
            this.Button5.Location = new System.Drawing.Point(204, 342);
            this.Button5.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Button5.Name = "Button5";
            this.Button5.Size = new System.Drawing.Size(220, 62);
            this.Button5.TabIndex = 75;
            this.Button5.UseVisualStyleBackColor = true;
            // 
            // Button4
            // 
            this.Button4.Location = new System.Drawing.Point(204, 266);
            this.Button4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Button4.Name = "Button4";
            this.Button4.Size = new System.Drawing.Size(220, 62);
            this.Button4.TabIndex = 74;
            this.Button4.UseVisualStyleBackColor = true;
            // 
            // Button3
            // 
            this.Button3.Location = new System.Drawing.Point(204, 192);
            this.Button3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Button3.Name = "Button3";
            this.Button3.Size = new System.Drawing.Size(220, 62);
            this.Button3.TabIndex = 73;
            this.Button3.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(488, 114);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(72, 20);
            this.label7.TabIndex = 71;
            this.label7.Text = "aaaaaaa";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(488, 67);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 20);
            this.label6.TabIndex = 70;
            this.label6.Text = "aaaaaaa";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(69, 160);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(72, 20);
            this.label5.TabIndex = 69;
            this.label5.Text = "aaaaaaa";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(69, 114);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(72, 20);
            this.label4.TabIndex = 68;
            this.label4.Text = "aaaaaaa";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(69, 67);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 20);
            this.label3.TabIndex = 67;
            this.label3.Text = "aaaaaaa";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(488, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(72, 20);
            this.label2.TabIndex = 66;
            this.label2.Text = "aaaaaaa";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(69, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 20);
            this.label1.TabIndex = 65;
            this.label1.Text = "aaaaaaa";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Location = new System.Drawing.Point(592, 109);
            this.dateTimePicker2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(175, 26);
            this.dateTimePicker2.TabIndex = 64;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(592, 62);
            this.dateTimePicker1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(175, 26);
            this.dateTimePicker1.TabIndex = 63;
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(193, 155);
            this.textBox5.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(175, 26);
            this.textBox5.TabIndex = 61;
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(193, 109);
            this.textBox4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(175, 26);
            this.textBox4.TabIndex = 60;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(193, 62);
            this.textBox3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(175, 26);
            this.textBox3.TabIndex = 59;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(592, 16);
            this.textBox2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(175, 26);
            this.textBox2.TabIndex = 58;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(193, 16);
            this.textBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(175, 26);
            this.textBox1.TabIndex = 57;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.label14);
            this.panel4.Controls.Add(this.BtnOpenSapGui);
            this.panel4.Controls.Add(this.label12);
            this.panel4.Controls.Add(this.label13);
            this.panel4.Controls.Add(this.txtPassword);
            this.panel4.Controls.Add(this.txtUser);
            this.panel4.Location = new System.Drawing.Point(103, 171);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(788, 420);
            this.panel4.TabIndex = 79;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(334, 22);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(0, 20);
            this.label14.TabIndex = 90;
            // 
            // BtnOpenSapGui
            // 
            this.BtnOpenSapGui.Location = new System.Drawing.Point(350, 186);
            this.BtnOpenSapGui.Name = "BtnOpenSapGui";
            this.BtnOpenSapGui.Size = new System.Drawing.Size(129, 40);
            this.BtnOpenSapGui.TabIndex = 89;
            this.BtnOpenSapGui.UseVisualStyleBackColor = true;
            this.BtnOpenSapGui.Click += new System.EventHandler(this.BtnOpenSapGui_Click);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(200, 107);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(72, 20);
            this.label12.TabIndex = 88;
            this.label12.Text = "aaaaaaa";
            this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(200, 64);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(72, 20);
            this.label13.TabIndex = 87;
            this.label13.Text = "aaaaaaa";
            this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(315, 104);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(198, 26);
            this.txtPassword.TabIndex = 83;
            // 
            // txtUser
            // 
            this.txtUser.Location = new System.Drawing.Point(315, 61);
            this.txtUser.Name = "txtUser";
            this.txtUser.Size = new System.Drawing.Size(198, 26);
            this.txtUser.TabIndex = 82;
            // 
            // textBox7
            // 
            this.textBox7.Location = new System.Drawing.Point(211, 99);
            this.textBox7.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(622, 26);
            this.textBox7.TabIndex = 62;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.Color.Red;
            this.label9.Location = new System.Drawing.Point(258, 132);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(0, 28);
            this.label9.TabIndex = 64;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.ForeColor = System.Drawing.Color.Red;
            this.label10.Location = new System.Drawing.Point(224, 134);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(0, 26);
            this.label10.TabIndex = 65;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.ForeColor = System.Drawing.Color.LimeGreen;
            this.label11.Location = new System.Drawing.Point(218, 134);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(0, 28);
            this.label11.TabIndex = 66;
            // 
            // BtnOpenSAP
            // 
            this.BtnOpenSAP.Location = new System.Drawing.Point(15, 3);
            this.BtnOpenSAP.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BtnOpenSAP.Name = "BtnOpenSAP";
            this.BtnOpenSAP.Size = new System.Drawing.Size(220, 34);
            this.BtnOpenSAP.TabIndex = 59;
            this.BtnOpenSAP.Text = "Ouvrir SAP GUI";
            this.BtnOpenSAP.UseVisualStyleBackColor = true;
            this.BtnOpenSAP.Click += new System.EventHandler(this.BtnOpenSAP_Click);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.BtnOpenSAP);
            this.panel3.Location = new System.Drawing.Point(641, 55);
            this.panel3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(250, 39);
            this.panel3.TabIndex = 61;
            // 
            // panel6
            // 
            this.panel6.Controls.Add(this.RadioButton2);
            this.panel6.Controls.Add(this.RadioButton1);
            this.panel6.Location = new System.Drawing.Point(792, 12);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(209, 38);
            this.panel6.TabIndex = 82;
            // 
            // RadioButton2
            // 
            this.RadioButton2.AutoSize = true;
            this.RadioButton2.Location = new System.Drawing.Point(103, 3);
            this.RadioButton2.Name = "RadioButton2";
            this.RadioButton2.Size = new System.Drawing.Size(86, 24);
            this.RadioButton2.TabIndex = 83;
            this.RadioButton2.TabStop = true;
            this.RadioButton2.Text = "English";
            this.RadioButton2.UseVisualStyleBackColor = true;
            this.RadioButton2.CheckedChanged += new System.EventHandler(this.RadioButton2_CheckedChanged);
            // 
            // RadioButton1
            // 
            this.RadioButton1.AutoSize = true;
            this.RadioButton1.Location = new System.Drawing.Point(2, 3);
            this.RadioButton1.Name = "RadioButton1";
            this.RadioButton1.Size = new System.Drawing.Size(95, 24);
            this.RadioButton1.TabIndex = 82;
            this.RadioButton1.Text = "Français";
            this.RadioButton1.UseVisualStyleBackColor = true;
            this.RadioButton1.CheckedChanged += new System.EventHandler(this.RadioButton1_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1013, 603);
            this.Controls.Add(this.panel6);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.textBox7);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel5);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Close Activities";
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        //private Panel panel2;
        private Panel panel1;
        private Button BtnGetSQL;
        private Panel panel2;
        private Button BtnLoadSQL;
        private Panel panel5;
        private Button Button5;
        private Button Button4;
        private Button Button3;
        private Label label7;
        private Label label6;
        private Label label5;
        private Label label4;
        private Label label3;
        private Label label2;
        private Label label1;
        private DateTimePicker dateTimePicker2;
        private DateTimePicker dateTimePicker1;
        private TextBox textBox5;
        private TextBox textBox4;
        private TextBox textBox3;
        private TextBox textBox2;
        private TextBox textBox1;
        private TextBox textBox7;
        private Label label9;
        private Label label10;
        private Label label11;
        private Panel panel4;
        private Button BtnOpenSapGui;
        private Label label12;
        private Label label13;
        private TextBox txtPassword;
        private TextBox txtUser;
        private Button BtnOpenSAP;
        private Panel panel3;
        private Label label14;
        private Panel panel6;
        private RadioButton RadioButton2;
        private RadioButton RadioButton1;
    }
}

