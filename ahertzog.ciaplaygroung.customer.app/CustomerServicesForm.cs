using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            
            SaveFileDialog saveFilexlsxDialog = new SaveFileDialog();
            saveFilexlsxDialog.Filter = filter;
            saveFilexlsxDialog.FilterIndex = 1;
            saveFilexlsxDialog.Title = "Salvar novo arquivo Excel";
            saveFilexlsxDialog.FileName = sourcePath.Split('/').Last()+"-Clientes";
            if (saveFilexlsxDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string sourceFile in Directory.GetFiles(sourcePath, "*.xls"))
                {
                    string fileName = Path.GetFileName(sourceFile);
                    customers.Add(customerServices.ReadFile(sourceFile));
                }
                try
                {
                    customerServices.WriteCustomersFile(customers, saveFilexlsxDialog.FileName);
                    MessageBox.Show($"Arquivo {saveFilexlsxDialog.FileName} criado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception)
                {
                    MessageBox.Show("Erro ao criar planilha com os cliente.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                
            }
        }

        private void ButtonSearch_Click(object sender, EventArgs e)
        {
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
