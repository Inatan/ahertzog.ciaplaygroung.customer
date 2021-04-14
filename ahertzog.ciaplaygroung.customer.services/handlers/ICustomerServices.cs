using ahertzog.ciaplaygroung.customer.domain.model;

namespace ahertzog.ciaplaygroung.customer.services.handlers
{
    public interface ICustomerServices
    {
        Customer ReadFile(string filePath);

        void WriteCustomersFile(CustomerServices customer);
    }
}
