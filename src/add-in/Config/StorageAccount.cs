using System.Collections.Specialized;
using System.Globalization;
using Microsoft.Win32;

namespace MyJournal.Notebook.Config
{
    class StorageAccount
    {
        private readonly NameValueCollection _cid;

        internal StorageAccount()
        {
            _cid = new NameValueCollection
            {
                { DEFAULT, null }
            };
            AddOfficeIDs();
            AddOneDriveIDs();
        }

        void AddOfficeIDs()
        {
            foreach (var version in s_office_version_list)
            {
                var subKey = string.Format(OFFICE_IDENTITIES_SUBKEY, version);
                using (var k = Registry.CurrentUser.OpenSubKey(subKey))
                {
                    if (k != null)
                    {
                        foreach (var identity in k.GetSubKeyNames())
                        {
                            using (var key = k.OpenSubKey(identity))
                            {
                                // Skip accounts that are in an error state
                                var errorState = (int)key.GetValue("ErrorState", 0);
                                if (errorState != 0) continue;

                                // Select Windows Live IDs
                                var idP = (int)key.GetValue("IdP", 0);
                                if (idP == Windows_Live_IdP)
                                {
                                    var email = (string)key.GetValue("EmailAddress");
                                    var id = email?.ToLower(CultureInfo.CurrentCulture);

                                    _cid[id] = (string)key.GetValue("ProviderId");
                                }
                            }
                        }
                        break;
                    }
                }
            }
        }

        void AddOneDriveIDs()
        {
            var subKey = @"SOFTWARE\Microsoft\IdentityCRL\UserExtendedProperties";
            using (var k = Registry.CurrentUser.OpenSubKey(subKey))
            {
                if (k != null)
                {
                    foreach (var msAccount in k.GetSubKeyNames())
                    {
                        using (var key = k.OpenSubKey(msAccount))
                        {
                            var id = msAccount.ToLower(CultureInfo.CurrentCulture);

                            _cid[id] = (string)key.GetValue("cid");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// An associative array to look up the OneDrive Customer ID (CID) for a
        /// Microsoft Account key.
        /// </summary>
        internal NameValueCollection CID => _cid;

        /// <summary>
        /// Returns true if the Journal notebook is stored on local storage;
        /// false if stored in the cloud.
        /// </summary>
        internal static bool IsDefault =>
            (Properties.Settings.Default.StorageAccount == DEFAULT);

        /// <summary>
        /// Retrieves an array of strings that contains all the storage account
        /// keys (Microsoft Accounts) for the current user.
        /// </summary>
        internal object[] Items => _cid.AllKeys;

        /// <summary>
        /// Default local storage account key.
        /// </summary>
        internal static string DEFAULT => "( local )";

        // Office 2016 and 2013 version numbers
        static readonly int[] s_office_version_list = { 16, 15 };

        const string OFFICE_IDENTITIES_SUBKEY =
            @"SOFTWARE\Microsoft\Office\{0}.0\Common\Identity\Identities";

        const int Windows_Live_IdP = 1;
    }
}
