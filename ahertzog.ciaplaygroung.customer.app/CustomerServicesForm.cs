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
            ButtonGenerate.Enabled = false;
        }

        private void ButtonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ButtonGenerate_Click(object sender, EventArgs e)
        {
            string sourcePath = textBoxPath.Text;
            List<Customer> customers = new List<Customer>();
            CustomerServices customerServices = new CustomerServices();
            foreach (string sourceFile in Directory.GetFiles(sourcePath, "*.xlsx"))
            {
                string fileName = Path.GetFileName(sourceFile);
                customers.Add(customerServices.ReadFile(sourceFile));
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
