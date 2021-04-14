using ahertzog.ciaplaygroung.customer.domain.model;
using System.Collections.Generic;

namespace ahertzog.ciaplaygroung.customer.services.handlers
{
    public interface ICustomerServices
    {
        Customer ReadFile(string filePath);

        void WriteCustomersFile(IList<Customer> customers,string newFilePath);
    }
}
