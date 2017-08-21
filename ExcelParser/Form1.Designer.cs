namespace ExcelParser
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
            this.pg_bar = new System.Windows.Forms.ProgressBar();
            this.start = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.l_lost = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.l_proc = new System.Windows.Forms.Label();
            this.l_nproc = new System.Windows.Forms.Label();
            this.l_all = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.l_cur_time = new System.Windows.Forms.Label();
            this.ofd = new System.Windows.Forms.OpenFileDialog();
            this.label6 = new System.Windows.Forms.Label();
            this.link_file = new System.Windows.Forms.LinkLabel();
            this.b_file = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // pg_bar
            // 
            this.pg_bar.Location = new System.Drawing.Point(12, 84);
            this.pg_bar.Name = "pg_bar";
            this.pg_bar.Size = new System.Drawing.Size(398, 23);
            this.pg_bar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.pg_bar.TabIndex = 0;
            // 
            // start
            // 
            this.start.Enabled = false;
            this.start.Location = new System.Drawing.Point(180, 200);
            this.start.Name = "start";
            this.start.Size = new System.Drawing.Size(75, 23);
            this.start.TabIndex = 1;
            this.start.Text = "Парсинг";
            this.start.UseVisualStyleBackColor = true;
            this.start.Click += new System.EventHandler(this.start_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 68);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Оставшееся время:";
            // 
            // l_lost
            // 
            this.l_lost.AutoSize = true;
            this.l_lost.Location = new System.Drawing.Point(126, 68);
            this.l_lost.Name = "l_lost";
            this.l_lost.Size = new System.Drawing.Size(49, 13);
            this.l_lost.TabIndex = 3;
            this.l_lost.Text = "00:00:00";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 133);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Обработано:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 156);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(86, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Не обработано:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 178);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(40, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Всего:";
            // 
            // l_proc
            // 
            this.l_proc.AutoSize = true;
            this.l_proc.Location = new System.Drawing.Point(99, 133);
            this.l_proc.Name = "l_proc";
            this.l_proc.Size = new System.Drawing.Size(13, 13);
            this.l_proc.TabIndex = 7;
            this.l_proc.Text = "0";
            // 
            // l_nproc
            // 
            this.l_nproc.AutoSize = true;
            this.l_nproc.Location = new System.Drawing.Point(99, 156);
            this.l_nproc.Name = "l_nproc";
            this.l_nproc.Size = new System.Drawing.Size(13, 13);
            this.l_nproc.TabIndex = 8;
            this.l_nproc.Text = "0";
            // 
            // l_all
            // 
            this.l_all.AutoSize = true;
            this.l_all.Location = new System.Drawing.Point(99, 178);
            this.l_all.Name = "l_all";
            this.l_all.Size = new System.Drawing.Size(13, 13);
            this.l_all.TabIndex = 9;
            this.l_all.Text = "0";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 110);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(110, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Затраченное время:";
            // 
            // l_cur_time
            // 
            this.l_cur_time.AutoSize = true;
            this.l_cur_time.Location = new System.Drawing.Point(126, 110);
            this.l_cur_time.Name = "l_cur_time";
            this.l_cur_time.Size = new System.Drawing.Size(49, 13);
            this.l_cur_time.TabIndex = 3;
            this.l_cur_time.Text = "00:00:00";
            // 
            // ofd
            // 
            this.ofd.Filter = "XLS|*.xls|XLSX|*.xlsx";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(15, 13);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(98, 13);
            this.label6.TabIndex = 11;
            this.label6.Text = "Выбранный файл:";
            // 
            // link_file
            // 
            this.link_file.AutoSize = true;
            this.link_file.LinkColor = System.Drawing.Color.Black;
            this.link_file.Location = new System.Drawing.Point(119, 13);
            this.link_file.Name = "link_file";
            this.link_file.Size = new System.Drawing.Size(16, 13);
            this.link_file.TabIndex = 13;
            this.link_file.TabStop = true;
            this.link_file.Text = "...";
            this.link_file.Click += new System.EventHandler(this.b_file_Click);
            // 
            // b_file
            // 
            this.b_file.Location = new System.Drawing.Point(335, 8);
            this.b_file.Name = "b_file";
            this.b_file.Size = new System.Drawing.Size(75, 23);
            this.b_file.TabIndex = 14;
            this.b_file.Text = "Выбрать";
            this.b_file.UseVisualStyleBackColor = true;
            this.b_file.Click += new System.EventHandler(this.b_file_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(422, 236);
            this.Controls.Add(this.b_file);
            this.Controls.Add(this.link_file);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.l_all);
            this.Controls.Add(this.l_nproc);
            this.Controls.Add(this.l_proc);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.l_cur_time);
            this.Controls.Add(this.l_lost);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.start);
            this.Controls.Add(this.pg_bar);
            this.Name = "Form1";
            this.Text = "ExcelParser";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar pg_bar;
        private System.Windows.Forms.Button start;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label l_lost;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label l_proc;
        private System.Windows.Forms.Label l_nproc;
        private System.Windows.Forms.Label l_all;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label l_cur_time;
        private System.Windows.Forms.OpenFileDialog ofd;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.LinkLabel link_file;
        private System.Windows.Forms.Button b_file;
    }
}

