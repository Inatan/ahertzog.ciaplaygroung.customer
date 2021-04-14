
namespace ahertzog.ciaplaygroung.customer.app
{
    partial class CustomerServicesForm
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBoxPath = new System.Windows.Forms.TextBox();
            this.ButtonGenerate = new System.Windows.Forms.Button();
            this.ButtonSearch = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.ButtonClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBoxPath
            // 
            this.textBoxPath.Enabled = false;
            this.textBoxPath.Location = new System.Drawing.Point(81, 37);
            this.textBoxPath.Name = "textBoxPath";
            this.textBoxPath.Size = new System.Drawing.Size(433, 20);
            this.textBoxPath.TabIndex = 0;
            // 
            // ButtonGenerate
            // 
            this.ButtonGenerate.Location = new System.Drawing.Point(8, 75);
            this.ButtonGenerate.Name = "ButtonGenerate";
            this.ButtonGenerate.Size = new System.Drawing.Size(149, 23);
            this.ButtonGenerate.TabIndex = 1;
            this.ButtonGenerate.Text = "Gerar Excel de Clientes";
            this.ButtonGenerate.UseVisualStyleBackColor = true;
            this.ButtonGenerate.Click += new System.EventHandler(this.ButtonGenerate_Click);
            // 
            // ButtonSearch
            // 
            this.ButtonSearch.Location = new System.Drawing.Point(520, 35);
            this.ButtonSearch.Name = "ButtonSearch";
            this.ButtonSearch.Size = new System.Drawing.Size(75, 23);
            this.ButtonSearch.TabIndex = 1;
            this.ButtonSearch.Text = "Procurar";
            this.ButtonSearch.UseVisualStyleBackColor = true;
            this.ButtonSearch.Click += new System.EventHandler(this.ButtonSearch_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Pasta Origem";
            // 
            // ButtonClose
            // 
            this.ButtonClose.Location = new System.Drawing.Point(520, 75);
            this.ButtonClose.Name = "ButtonClose";
            this.ButtonClose.Size = new System.Drawing.Size(75, 23);
            this.ButtonClose.TabIndex = 1;
            this.ButtonClose.Text = "Fechar";
            this.ButtonClose.UseVisualStyleBackColor = true;
            this.ButtonClose.Click += new System.EventHandler(this.ButtonClose_Click);
            // 
            // CustomerServicesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(615, 110);
            this.ControlBox = false;
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ButtonSearch);
            this.Controls.Add(this.ButtonClose);
            this.Controls.Add(this.ButtonGenerate);
            this.Controls.Add(this.textBoxPath);
            this.Name = "CustomerServicesForm";
            this.Text = "Conversor de Clientes";
            this.Load += new System.EventHandler(this.FormCustomer_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxPath;
        private System.Windows.Forms.Button ButtonGenerate;
        private System.Windows.Forms.Button ButtonSearch;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button ButtonClose;
    }
}

