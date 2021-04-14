using System.Collections.Generic;

namespace ahertzog.ciaplaygroung.customer.domain.model
{
    public class Customer
    {
        public Customer()
        {
            Emails = new List<string>();
        }

        public string CompanyName { get; set; }
        public string FantasyName { get; set; }
        public string Address { get; set; }
        public string District { get; set; }
        public string CEP { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string CNPJ { get; set; }
        public string IE { get; set; }
        public string Phone { get; set; }
        public string CellularNumber { get; set; }
        public string Contact { get; set; }
        public IList<string> Emails { get; set; }

    }
}
