//Copyright (c) 2008 Lenel Systems International, Inc.  

using System;
using System.Management;
using System.Threading;
using System.Windows.Forms;


namespace LoggedEventsSample
{
    /// <summary>
    /// The sample demostrates how to get all the logged events for panelID=1, using DataConduIT.
    /// </summary>
    public class GetLoggedEvents
    {
        public static void Main()
        {
            try
            {
                // Get list of all the logged events for panel ID=117. The object is initiated with the path
                // and the query.
                ManagementObjectSearcher searcher = 
                    new ManagementObjectSearcher("root\\OnGuard",
                    "SELECT * FROM Lnl_LoggedEvent where PanelID=117"); 
                
                Console.WriteLine("Lnl_LoggedEvent instance - Panel ID 117");
                foreach (ManagementObject queryObj in searcher.Get())
                {
                    Console.WriteLine("-----------------------------------");
                    Console.WriteLine("Description: {0}", queryObj["Description"]);
                    Console.WriteLine("Type: {0}", queryObj["Type"]);
                    Console.WriteLine("Serial Number: {0}", queryObj["SerialNumber"]);
                    Console.WriteLine("DeviceID: {0}", queryObj["DeviceID"]);
                    Console.WriteLine("Time: {0}", queryObj["Time"]);
                    Thread.Sleep(1500);

                }
                Thread.Sleep(10000);
            }
            catch (ManagementException e)
            {
                MessageBox.Show("An error occurred while querying for WMI data: " + e.Message);
            }
        }
    }
}
