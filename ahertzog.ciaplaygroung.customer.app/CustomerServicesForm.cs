using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using ahertzog.ciaplaygroung.customer.domain.model;
using ahertzog.ciaplaygroung.customer.services.handlers;

namespace ahertzog.ciaplaygroung.customer.app
{
    public partial class CustomerServicesForm : Form
    {
        private readonly ICustomerServices customerServices;
        public CustomerServicesForm()
        {
            InitializeComponent();
            customerServices = new CustomerServices();
            LabelProgress.Visible = false;
            ProgressBarCustomers.Visible = false;
            LabelProgress.Text = string.Empty;

        }

        private void FormCustomer_Load(object sender, EventArgs e)
        {
            ButtonGenerate.Enabled = false;
        }

        private void ButtonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ButtonGenerate_Click(object sender, EventArgs e)
        {
            string filter = "Excel Worksheets 2007(*.xlsx)|*.xlsx";
            string sourcePath = textBoxPath.Text;
            List<Customer> customers = new List<Customer>();

            using (SaveFileDialog saveFilexlsxDialog = new SaveFileDialog())
            {
                saveFilexlsxDialog.Filter = filter;
                saveFilexlsxDialog.FilterIndex = 1;
                saveFilexlsxDialog.Title = "Salvar novo arquivo Excel";
                saveFilexlsxDialog.FileName = sourcePath.Split('/').Last()+"-Clientes";
                if (saveFilexlsxDialog.ShowDialog() == DialogResult.OK)
                {
                    var files = Directory.GetFiles(sourcePath, "*.xls");
                    if (files.Length > 0)
                    {
                        LabelProgress.Visible = true;
                        LabelProgress.Text = $"Lendo arquivos de {saveFilexlsxDialog.FileName}...";
                        ProgressBarCustomers.Visible = true;

                        ProgressBarCustomers.Minimum = 0;
                        ProgressBarCustomers.Maximum = files.Length + 2;
                        ProgressBarCustomers.Step = 1;
                        try
                        {
                            foreach (string sourceFile in files)
                            {
                                ProgressBarCustomers.PerformStep();
                                customers.Add(customerServices.ReadFile(sourceFile));
                            }
                            LabelProgress.Text = $"Gerando arquivo {saveFilexlsxDialog.FileName}...";
                            ProgressBarCustomers.PerformStep();

                            customerServices.WriteCustomersFile(customers, saveFilexlsxDialog.FileName);
                            
                            ProgressBarCustomers.PerformStep();
                            LabelProgress.Text = $"Finalizado";

                            MessageBox.Show($"Arquivo {saveFilexlsxDialog.FileName} criado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception)
                        {
                            LabelProgress.Visible = false;
                            ProgressBarCustomers.Visible = false;
                            LabelProgress.Text = string.Empty;
                            MessageBox.Show("Erro ao criar planilha com os cliente.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        LabelProgress.Visible = false;
                        ProgressBarCustomers.Visible = false;
                        LabelProgress.Text = string.Empty;
                        MessageBox.Show($"O caminho selecionado {sourcePath}, não apresenta nenhum arquivo excel. Favor selecionar outro caminho.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private void ButtonSearch_Click(object sender, EventArgs e)
        {
            LabelProgress.Visible = false;
            ProgressBarCustomers.Visible = false;
            LabelProgress.Text = string.Empty;
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    textBoxPath.Text = fbd.SelectedPath;
                    string[] files = Directory.GetFiles(fbd.SelectedPath);
                    ButtonGenerate.Enabled = true;
                }
            }
        }
    }
}
