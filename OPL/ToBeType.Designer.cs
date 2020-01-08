namespace OPL
{
    partial class ToBeType
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
            this.NewFileName = new DevExpress.XtraEditors.TextEdit();
            this.LabNewFileName = new DevExpress.XtraEditors.LabelControl();
            this.NewFileCanel = new DevExpress.XtraEditors.SimpleButton();
            this.NewFileOK = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.NewFileName.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // NewFileName
            // 
            this.NewFileName.Location = new System.Drawing.Point(18, 21);
            this.NewFileName.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.NewFileName.Name = "NewFileName";
            this.NewFileName.Properties.Appearance.Font = new System.Drawing.Font("STXihei", 12F);
            this.NewFileName.Properties.Appearance.Options.UseFont = true;
            this.NewFileName.Size = new System.Drawing.Size(231, 28);
            this.NewFileName.TabIndex = 0;
            this.NewFileName.TextChanged += new System.EventHandler(this.NewFileName_TextChanged);
            this.NewFileName.Click += new System.EventHandler(this.NewFileName_Click);
            this.NewFileName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.NewFileName_KeyPress);
            // 
            // LabNewFileName
            // 
            this.LabNewFileName.Appearance.BackColor = System.Drawing.Color.White;
            this.LabNewFileName.Appearance.Font = new System.Drawing.Font("STXihei", 12F);
            this.LabNewFileName.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.LabNewFileName.Location = new System.Drawing.Point(23, 25);
            this.LabNewFileName.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.LabNewFileName.Name = "LabNewFileName";
            this.LabNewFileName.Size = new System.Drawing.Size(140, 21);
            this.LabNewFileName.TabIndex = 2;
            this.LabNewFileName.Text = "请输入新文件名";
            this.LabNewFileName.Click += new System.EventHandler(this.NewFileName_Click);
            // 
            // NewFileCanel
            // 
            this.NewFileCanel.Appearance.Font = new System.Drawing.Font("STXihei", 12F);
            this.NewFileCanel.Appearance.Options.UseFont = true;
            this.NewFileCanel.Location = new System.Drawing.Point(170, 59);
            this.NewFileCanel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.NewFileCanel.Name = "NewFileCanel";
            this.NewFileCanel.Size = new System.Drawing.Size(79, 33);
            this.NewFileCanel.TabIndex = 3;
            this.NewFileCanel.Text = "Cancel";
            this.NewFileCanel.Click += new System.EventHandler(this.NewFileCanel_Click);
            // 
            // NewFileOK
            // 
            this.NewFileOK.Appearance.Font = new System.Drawing.Font("STXihei", 12F);
            this.NewFileOK.Appearance.Options.UseFont = true;
            this.NewFileOK.Location = new System.Drawing.Point(22, 59);
            this.NewFileOK.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.NewFileOK.Name = "NewFileOK";
            this.NewFileOK.Size = new System.Drawing.Size(79, 33);
            this.NewFileOK.TabIndex = 4;
            this.NewFileOK.Text = "OK";
            this.NewFileOK.Click += new System.EventHandler(this.NewFileOK_Click);
            // 
            // ToBeType
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(286, 120);
            this.ControlBox = false;
            this.Controls.Add(this.NewFileOK);
            this.Controls.Add(this.NewFileCanel);
            this.Controls.Add(this.LabNewFileName);
            this.Controls.Add(this.NewFileName);
            this.FormBorderEffect = DevExpress.XtraEditors.FormBorderEffect.Glow;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "ToBeType";
            this.Text = "ToBeType";
            this.Load += new System.EventHandler(this.ToBeType_Load);
            ((System.ComponentModel.ISupportInitialize)(this.NewFileName.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraEditors.TextEdit NewFileName;
        private DevExpress.XtraEditors.LabelControl LabNewFileName;
        private DevExpress.XtraEditors.SimpleButton NewFileCanel;
        private DevExpress.XtraEditors.SimpleButton NewFileOK;
    }
}