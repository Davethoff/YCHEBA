
namespace Client
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.textBoxLocalPort = new System.Windows.Forms.TextBox();
            this.buttonBind = new System.Windows.Forms.Button();
            this.textBoxLog = new System.Windows.Forms.TextBox();
            this.listBoxClients = new System.Windows.Forms.ListBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.vScrollBar1 = new System.Windows.Forms.VScrollBar();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // textBoxLocalPort
            // 
            this.textBoxLocalPort.Location = new System.Drawing.Point(190, 12);
            this.textBoxLocalPort.Name = "textBoxLocalPort";
            this.textBoxLocalPort.Size = new System.Drawing.Size(100, 20);
            this.textBoxLocalPort.TabIndex = 0;
            // 
            // buttonBind
            // 
            this.buttonBind.Location = new System.Drawing.Point(298, 9);
            this.buttonBind.Name = "buttonBind";
            this.buttonBind.Size = new System.Drawing.Size(96, 23);
            this.buttonBind.TabIndex = 1;
            this.buttonBind.Text = "Открыть Порт";
            this.buttonBind.UseVisualStyleBackColor = true;
            this.buttonBind.Click += new System.EventHandler(this.buttonBind_Click);
            // 
            // textBoxLog
            // 
            this.textBoxLog.Location = new System.Drawing.Point(4, 38);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.Size = new System.Drawing.Size(475, 407);
            this.textBoxLog.TabIndex = 2;
            // 
            // listBoxClients
            // 
            this.listBoxClients.FormattingEnabled = true;
            this.listBoxClients.Location = new System.Drawing.Point(485, 38);
            this.listBoxClients.Name = "listBoxClients";
            this.listBoxClients.Size = new System.Drawing.Size(303, 407);
            this.listBoxClients.TabIndex = 3;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // vScrollBar1
            // 
            this.vScrollBar1.Location = new System.Drawing.Point(461, 38);
            this.vScrollBar1.Name = "vScrollBar1";
            this.vScrollBar1.Size = new System.Drawing.Size(17, 407);
            this.vScrollBar1.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(93, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(91, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Локальный порт";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.vScrollBar1);
            this.Controls.Add(this.listBoxClients);
            this.Controls.Add(this.textBoxLog);
            this.Controls.Add(this.buttonBind);
            this.Controls.Add(this.textBoxLocalPort);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxLocalPort;
        private System.Windows.Forms.Button buttonBind;
        private System.Windows.Forms.TextBox textBoxLog;
        private System.Windows.Forms.ListBox listBoxClients;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.VScrollBar vScrollBar1;
        private System.Windows.Forms.Label label1;
    }
}

