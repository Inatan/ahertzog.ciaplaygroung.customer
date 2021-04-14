using ahertzog.ciaplaygroung.customer.domain.model;
using ExcelDataReader;
using System;
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
                        if(reader.GetString(0) != null)
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
                    }
                }
            }
            return customer;
        }

        public void WriteCustomersFile(CustomerServices customer)
        {
            throw new NotImplementedException();
        }
    }
}
