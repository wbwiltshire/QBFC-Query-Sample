/*************************************************************************************************/
//
// Link: https://github.com/Schparky/qbfc-query
//
/************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QBFC13Lib;

namespace QBFC_Query_Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            ListCustomers();
        }

        static public void ListCustomers()
        {
            QBSession Session = null;

            try
            {
                Session = new QBSession();

                IResponse response = Session.QueryCustomerList();

                ICustomerRetList customerRetList = response.Detail as ICustomerRetList;
                for (var i = 0; i < customerRetList.Count; i++)
                {
                    ICustomerRet customerRet = customerRetList.GetAt(i);
                    Console.WriteLine(customerRet.FullName.GetValue());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
                Console.WriteLine("Stack Trace:");
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine("Exiting the application");
            }
        }
    }
}
