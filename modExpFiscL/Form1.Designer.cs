namespace modExpFiscL
{
    partial class FrmExport
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.comboBoxBase = new System.Windows.Forms.ComboBox();
            this.textBoxServeur = new System.Windows.Forms.TextBox();
            this.textBoxPort = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnExtraire = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.comboBoxBase);
            this.groupBox1.Controls.Add(this.textBoxServeur);
            this.groupBox1.Controls.Add(this.textBoxPort);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(360, 109);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Sélectionner le serveur et la base de données pour laquelle réaliser l\'extraction" +
    " :";
            // 
            // comboBoxBase
            // 
            this.comboBoxBase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxBase.FormattingEnabled = true;
            this.comboBoxBase.Location = new System.Drawing.Point(9, 76);
            this.comboBoxBase.Name = "comboBoxBase";
            this.comboBoxBase.Size = new System.Drawing.Size(345, 21);
            this.comboBoxBase.TabIndex = 4;
            // 
            // textBoxServeur
            // 
            this.textBoxServeur.Location = new System.Drawing.Point(247, 41);
            this.textBoxServeur.Name = "textBoxServeur";
            this.textBoxServeur.Size = new System.Drawing.Size(107, 20);
            this.textBoxServeur.TabIndex = 3;
            this.textBoxServeur.Text = "127.0.0.1";
            this.textBoxServeur.KeyUp += new System.Windows.Forms.KeyEventHandler(this.textBoxServeur_KeyUp);
            // 
            // textBoxPort
            // 
            this.textBoxPort.Location = new System.Drawing.Point(44, 41);
            this.textBoxPort.Name = "textBoxPort";
            this.textBoxPort.Size = new System.Drawing.Size(40, 20);
            this.textBoxPort.TabIndex = 2;
            this.textBoxPort.Text = "5432";
            this.textBoxPort.KeyUp += new System.Windows.Forms.KeyEventHandler(this.textBoxPort_KeyUp);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(104, 44);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(137, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Adresse / Nom du serveur :";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 44);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Port :";
            // 
            // btnExtraire
            // 
            this.btnExtraire.Location = new System.Drawing.Point(155, 127);
            this.btnExtraire.Name = "btnExtraire";
            this.btnExtraire.Size = new System.Drawing.Size(75, 23);
            this.btnExtraire.TabIndex = 1;
            this.btnExtraire.Text = "Extraire";
            this.btnExtraire.UseVisualStyleBackColor = true;
            this.btnExtraire.Click += new System.EventHandler(this.btnExtraire_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 163);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(360, 23);
            this.progressBar1.TabIndex = 2;
            // 
            // FrmExport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(382, 198);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnExtraire);
            this.Controls.Add(this.groupBox1);
            this.Name = "FrmExport";
            this.Text = "Exportation des données ficscales";
            this.Load += new System.EventHandler(this.FrmExport_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox comboBoxBase;
        private System.Windows.Forms.TextBox textBoxServeur;
        private System.Windows.Forms.TextBox textBoxPort;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnExtraire;
        private System.Windows.Forms.ProgressBar progressBar1;
    }
}

