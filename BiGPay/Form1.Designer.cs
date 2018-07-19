namespace BiGPay
{
    partial class Form
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
            this.components = new System.ComponentModel.Container();
            this.SelectDossier = new System.Windows.Forms.FolderBrowserDialog();
            this.Dossier = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.timer = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // Dossier
            // 
            this.Dossier.BackgroundImage = global::BiGPay.Properties.Resources.folder__1_;
            this.Dossier.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.Dossier.Cursor = System.Windows.Forms.Cursors.Default;
            this.Dossier.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Dossier.ForeColor = System.Drawing.SystemColors.ControlLight;
            this.Dossier.Location = new System.Drawing.Point(53, 24);
            this.Dossier.Name = "Dossier";
            this.Dossier.Size = new System.Drawing.Size(122, 123);
            this.Dossier.TabIndex = 0;
            this.Dossier.UseVisualStyleBackColor = true;
            this.Dossier.Click += new System.EventHandler(this.Dossier_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 163);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(218, 12);
            this.progressBar.TabIndex = 2;
            // 
            // Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Maroon;
            this.ClientSize = new System.Drawing.Size(242, 196);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.Dossier);
            this.Name = "Form";
            this.Text = "BiGPay";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog SelectDossier;
        private System.Windows.Forms.Button Dossier;
        public System.Windows.Forms.Timer timer;
        public System.Windows.Forms.ProgressBar progressBar;
    }
}

