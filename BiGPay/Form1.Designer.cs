namespace BiGPay
{
    partial class Form1
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
            this.SelectDossier = new System.Windows.Forms.FolderBrowserDialog();
            this.Dossier = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Dossier
            // 
            this.Dossier.Location = new System.Drawing.Point(59, 49);
            this.Dossier.Name = "Dossier";
            this.Dossier.Size = new System.Drawing.Size(116, 48);
            this.Dossier.TabIndex = 0;
            this.Dossier.Text = "Sélectionner un dossier";
            this.Dossier.UseVisualStyleBackColor = true;
            this.Dossier.Click += new System.EventHandler(this.Dossier_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(242, 196);
            this.Controls.Add(this.Dossier);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog SelectDossier;
        private System.Windows.Forms.Button Dossier;
    }
}

