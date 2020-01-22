using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QBFC13Lib;

namespace QBFC_Query_Sample
{
    public class QBSession
    {

        private QBSessionManager SessionManager;

        private IMsgSetRequest requestMsgSet;

        /// Announcing QuickBooks version
        public double QBVer;

        public QBSession()
        {
            SessionManager = new QBSessionManager();

            SessionManager.OpenConnection("", "QBFC Query Tool");

            // use omDontCare rather than omSingleUser or omMultiUser
            SessionManager.BeginSession("", ENOpenMode.omDontCare);

            QBVer = QBFCLatestVersion();
            Console.WriteLine("The qbXML version v" + Convert.ToString(QBVer) + ".0 is detected. Applicaton will set its " +
                "compatibility accordingly.\n\nThis sample uses QBFC for all of its communication to QuickBooks.");

            // Get a RequestMsgSet based on the correct QB Version
            getLatestMsgSetRequest();
        }

        ~QBSession()
        {
            SessionManager.EndSession();
            SessionManager.CloseConnection();
        }

        public IResponse QueryCustomerList()
        {
            // Add the request to the message set request object
            ICustomerQuery CustQ = requestMsgSet.AppendCustomerQueryRq();

            // Optionally, you can put filter on it.
            CustQ.ORCustomerListQuery.CustomerListFilter.MaxReturned.SetValue(5);

            // Do the request and get the response message set object
            IMsgSetResponse responseSet = SessionManager.DoRequests(requestMsgSet);

            // Uncomment the following to view and save the request and response XML
            //string requestXML = requestSet.ToXMLString();
            //Console.WriteLine(requestXML);
            //string responseXML = responseSet.ToXMLString();
            //Console.WriteLine(responseXML);

            IResponse response = responseSet.ResponseList.GetAt(0);
            //int statusCode = response.StatusCode;
            //string statusMessage = response.StatusMessage;
            //string statusSeverity = response.StatusSeverity;
            //Console.WriteLine("Status:\nCode = " + statusCode + "\nMessage = " + statusMessage + "\nSeverity = " + statusSeverity);

            return response;
        }

        /// <summary>
        /// Get a MsgSetRequest for latest supported version of QBFC.
        /// </summary>
        private void getLatestMsgSetRequest()
        {
            short qbXMLMajorVer = 0;
            short qbXMLMinorVer = 0;

            if (QBVer >= 6.0)
            {
                qbXMLMajorVer = 6;
                qbXMLMinorVer = 0;
            }
            else if (QBVer >= 5.0)
            {
                qbXMLMajorVer = 5;
                qbXMLMinorVer = 0;
            }
            else if (QBVer >= 4.0)
            {
                qbXMLMajorVer = 4;
                qbXMLMinorVer = 0;
            }
            else if (QBVer >= 3.0)
            {
                qbXMLMajorVer = 3;
                qbXMLMinorVer = 0;
            }
            else if (QBVer >= 2.0)
            {
                qbXMLMajorVer = 2;
                qbXMLMinorVer = 0;
            }
            else if (QBVer >= 1.1)
            {
                qbXMLMajorVer = 1;
                qbXMLMinorVer = 1;
            }
            else
            {
                qbXMLMajorVer = 1;
                qbXMLMinorVer = 0;
                Console.WriteLine("It seems that you are running QuickBooks 2002 Release 1. We strongly " +
                    "recommend that you use QuickBooks' online update feature to obtain the latest fixes " +
                    "and enhancements");
            }

            // Create the message set request object
            requestMsgSet = SessionManager.CreateMsgSetRequest("US", qbXMLMajorVer, qbXMLMinorVer);

            // Initialize the message set request object
            requestMsgSet.Attributes.OnError = ENRqOnError.roeStop;
        }

        /// <summary>
        /// Method for handling different versions of QuickBooks.
        /// </summary>
        private double QBFCLatestVersion()
        {
            // Use oldest version to ensure that this application work with any QuickBooks (US)
            IMsgSetRequest msgset = SessionManager.CreateMsgSetRequest("US", 1, 0);
            msgset.AppendHostQueryRq();
            IMsgSetResponse QueryResponse = SessionManager.DoRequests(msgset);
            //Console.WriteLine("Host query = " + msgset.ToXMLString());

            // The response list contains only one response,
            // which corresponds to our single HostQuery request.
            IResponse response = QueryResponse.ResponseList.GetAt(0);

            // Please refer to QBFC Developers Guide for details on why 
            // "as" clause was used to link this derived class to its base class.
            IHostRet HostResponse = response.Detail as IHostRet;
            IBSTRList supportedVersions = HostResponse.SupportedQBXMLVersionList as IBSTRList;

            double LastVers = 0;

            for (var i = 0; i <= supportedVersions.Count - 1; i++)
            {
                string svers = null;
                svers = supportedVersions.GetAt(i);
                var vers = Convert.ToDouble(svers);
                if (vers > LastVers)
                {
                    LastVers = vers;
                }
            }
            return LastVers;
        }
    }
}
