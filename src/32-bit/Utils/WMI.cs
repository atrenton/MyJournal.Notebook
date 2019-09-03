using System.Management;

namespace MyJournal.Notebook.Utils
{
    /// <summary>
    /// Implements Windows Management Instrumentation (WMI) helper methods.
    /// </summary>
    static class WMI
    {
        /// <summary>
        /// Gets a list of active local user accounts that are not locked out.
        /// </summary>
        /// <returns>A collection of local user accounts</returns>
        internal static ManagementObjectCollection GetLocalUserAccounts()
        {
            const string WMI_Query = @"
                SELECT * FROM Win32_UserAccount
                 WHERE LocalAccount = TRUE
                   AND AccountType = 512
                   AND Disabled = FALSE
                   AND Lockout = FALSE
                   AND SIDType = 1
                   AND Status = 'OK'";
            var scope = new ManagementScope(@"\\.\Root\CIMv2");
            var query = new ObjectQuery(WMI_Query);

            ManagementObjectCollection collection = null;
            using (var searcher = new ManagementObjectSearcher(scope, query))
            {
                collection = searcher.Get();
            }

            return collection;
        }
    }
}
