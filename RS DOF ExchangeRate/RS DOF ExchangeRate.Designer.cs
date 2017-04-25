namespace RS_DOF_ExchangeRate
{
    partial class RS_Main
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RS_Main));
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblDate = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblLastUpdate = new System.Windows.Forms.Label();
            this.btnForceUpdate = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // timer1
            // 
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.RoyalBlue;
            this.label1.Location = new System.Drawing.Point(1, 97);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(350, 22);
            this.label1.TabIndex = 0;
            this.label1.Text = "Update DOF Exchange Rate Automation";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(33, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Date:";
            // 
            // lblDate
            // 
            this.lblDate.AutoSize = true;
            this.lblDate.Location = new System.Drawing.Point(58, 9);
            this.lblDate.Name = "lblDate";
            this.lblDate.Size = new System.Drawing.Size(40, 13);
            this.lblDate.TabIndex = 2;
            this.lblDate.Text = "lblDate";
            this.lblDate.Visible = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 51);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(106, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "Last Update Applied:";
            // 
            // lblLastUpdate
            // 
            this.lblLastUpdate.AutoSize = true;
            this.lblLastUpdate.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLastUpdate.ForeColor = System.Drawing.Color.DodgerBlue;
            this.lblLastUpdate.Location = new System.Drawing.Point(124, 46);
            this.lblLastUpdate.Name = "lblLastUpdate";
            this.lblLastUpdate.Size = new System.Drawing.Size(104, 18);
            this.lblLastUpdate.TabIndex = 4;
            this.lblLastUpdate.Text = "lblLastUodate";
            this.lblLastUpdate.Visible = false;
            // 
            // btnForceUpdate
            // 
            this.btnForceUpdate.Location = new System.Drawing.Point(234, 8);
            this.btnForceUpdate.Name = "btnForceUpdate";
            this.btnForceUpdate.Size = new System.Drawing.Size(103, 24);
            this.btnForceUpdate.TabIndex = 5;
            this.btnForceUpdate.Text = "Force Update";
            this.btnForceUpdate.UseVisualStyleBackColor = true;
            this.btnForceUpdate.Click += new System.EventHandler(this.btnForceUpdate_Click);
            // 
            // RS_Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(352, 130);
            this.Controls.Add(this.btnForceUpdate);
            this.Controls.Add(this.lblLastUpdate);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lblDate);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "RS_Main";
            this.Text = "FaMF - Exchange Rate Update - PLD System";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.RS_Main_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblDate;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblLastUpdate;
        private System.Windows.Forms.Button btnForceUpdate;
    }
}

