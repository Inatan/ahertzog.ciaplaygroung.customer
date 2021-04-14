using ahertzog.ciaplaygroung.customer.domain.model;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;

namespace ahertzog.ciaplaygroung.customer.services.handlers
{
    public class CustomerServices : ICustomerServices
    {
        public Customer ReadFile(string filePath)
        {
            Customer customer = new Customer();
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read())
                    {
                        if(reader.GetFieldType(0) == typeof(string) && reader.GetString(0) != null && reader.GetFieldType(1) == typeof(string))
                        {
                            string key = reader.GetString(0).Trim();
                            string value = reader.GetString(1);
                            switch (key)
                            {
                                case "RAZÃO SOCIAL":
                                    customer.CompanyName = value;
                                    break;
                                case "NOME FANTASIA":
                                    customer.FantasyName = value;
                                    break;
                                case "ENDEREÇO":
                                    customer.Address = value;
                                    break;
                                case "BAIRRO":
                                    customer.District = value;
                                    break;
                                case "CEP":
                                    customer.CEP = value;
                                    break;
                                case "CIDADE":
                                    customer.City = value;
                                    break;
                                case "ESTADO":
                                    customer.State = value;
                                    break;
                                case "CNPJ":
                                    customer.CNPJ = value;
                                    break;
                                case "I.E.":
                                    customer.IE = value;
                                    break;
                                case "TELEFONE":
                                    customer.Phone = value;
                                    break;
                                case "CELULAR":
                                    customer.CellularNumber = value;
                                    break;
                                case "CONTATO":
                                    customer.Contact = value;
                                    break;
                                case "E-MAIL":
                                    customer.Emails.Add(value);
                                    break;

                                default:
                                    break;
                            }
                        }
                        else if(reader.GetFieldType(0) == typeof(string) && reader.GetString(0) == null && reader.GetFieldType(1) == typeof(string) && reader.GetString(1) != null)
                        {
                            string value = reader.GetString(1);
                            if(value.Contains("@"))
                                customer.Emails.Add(value);
                        }
                    }
                }
            }
            return customer;
        }

        public void WriteCustomersFile(IList<Customer> customers, string newFilePath)
        {
            Application xlApp = new Application();
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
            
            xlWorkSheet.Cells[1, 1] = "RAZÃO SOCIAL";
            xlWorkSheet.Cells[1, 2] = "NOME FANTASIA";
            xlWorkSheet.Cells[1, 3] = "ENDEREÇO";
            xlWorkSheet.Cells[1, 4] = "BAIRRO";
            xlWorkSheet.Cells[1, 5] = "CEP";
            xlWorkSheet.Cells[1, 6] = "CIDADE";
            xlWorkSheet.Cells[1, 7] = "ESTADO";
            xlWorkSheet.Cells[1, 8] = "I.E.";
            xlWorkSheet.Cells[1, 9] = "TELEFONE";
            xlWorkSheet.Cells[1, 10] = "CELULAR";
            xlWorkSheet.Cells[1, 11] = "CONTATO";
            xlWorkSheet.Cells[1, 12] = "E-MAIL1";
            xlWorkSheet.Cells[1, 13] = "E-MAIL2";
            xlWorkSheet.Cells[1, 14] = "E-MAIL3";
            xlWorkSheet.get_Range("A1", "N1").Cells.Font.Bold = true;
            int row = 2;
            foreach (var customer in customers)
            {
                xlWorkSheet.Cells[row, 1] = customer.CompanyName;
                xlWorkSheet.Cells[row, 2] = customer.FantasyName;
                xlWorkSheet.Cells[row, 3] = customer.Address;
                xlWorkSheet.Cells[row, 4] = customer.District;
                xlWorkSheet.Cells[row, 5] = customer.CEP;
                xlWorkSheet.Cells[row, 6] = customer.City;
                xlWorkSheet.Cells[row, 7] = customer.State;
                xlWorkSheet.Cells[row, 8] = customer.IE;
                xlWorkSheet.Cells[row, 9] = customer.Phone;
                xlWorkSheet.Cells[row, 10] = customer.CellularNumber;
                xlWorkSheet.Cells[row, 11] = customer.Contact;
                for (int i = 0; i < customer.Emails.Count; i++)
                {
                    xlWorkSheet.Cells[row, 12+i] = customer.Emails[i];
                }

                row++;
            }
            xlWorkSheet.Columns.AutoFit();
            xlWorkBook.SaveAs(newFilePath, XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
        }
    }
}
