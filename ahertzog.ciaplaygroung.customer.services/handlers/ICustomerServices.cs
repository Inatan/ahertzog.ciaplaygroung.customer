namespace ahertzog.ciaplaygroung.customer.services.handlers
{
    interface ICustomerServices
    {
        Customer ReadFile();

        void WriteCustumersFile(Customer customer);
    }
}
