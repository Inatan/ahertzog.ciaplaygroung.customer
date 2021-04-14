using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using ahertzog.ciaplaygroung.customer.domain.model;
using ahertzog.ciaplaygroung.customer.services.handlers;

namespace ahertzog.ciaplaygroung.customer.app
{
    public partial class CustomerServicesForm : Form
    {
        public CustomerServicesForm()
        {
            InitializeComponent();
        }

        private void FormCustomer_Load(object sender, EventArgs e)
        {
            buttonGerar.Enabled = false;
        }

        private void ButtonFechar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ButtonGerar_Click(object sender, EventArgs e)
        {
            string sourcePath = textBoxPath.Text;
            List<Customer> customers = new List<Customer>();
            CustomerServices customerServices = new CustomerServices();
            foreach (string sourceFile in Directory.GetFiles(sourcePath, "*.xlsx"))
            {
                string fileName = Path.GetFileName(sourceFile);
                customerServices.ReadFile(sourceFile);
                //MessageBox.Show(fileName);
            }
        }

        private void ButtonProcurar_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    textBoxPath.Text = fbd.SelectedPath;
                    string[] files = Directory.GetFiles(fbd.SelectedPath);
                    buttonGerar.Enabled = true;
                }
            }
        }
    }
}
