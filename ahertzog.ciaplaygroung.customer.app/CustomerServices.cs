using System;
using System.IO;
using System.Windows.Forms;

namespace ahertzog.ciaplaygroung.customer.app
{
    public partial class CustomerServices : Form
    {
        public CustomerServices()
        {
            InitializeComponent();
        }

        private void FormCustomer_Load(object sender, EventArgs e)
        {

        }

        private void ButtonFechar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ButtonGerar_Click(object sender, EventArgs e)
        {

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

                    MessageBox.Show("Files found: " + files.Length.ToString(), "Message");
                }
            }
        }
    }
}
