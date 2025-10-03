using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.DirectoryServices;
using AccessManager.Classes;

namespace AccessManager.Services
{
    public class ActiveDirectoryHelper
    {
        private readonly DirectoryEntry _rootEntry;

        public ActiveDirectoryHelper(string ldapPath, string username = null, string password = null)
        {
            _rootEntry = string.IsNullOrEmpty(username)
                ? new DirectoryEntry(ldapPath)
                : new DirectoryEntry(ldapPath, username, password);
        }

        public List<AdObjectInfo> GetResources(string objectClass = "group")
        {
            var list = new List<AdObjectInfo>();

            using (var searcher = new DirectorySearcher(_rootEntry))
            {
                searcher.Filter = $"(objectClass={objectClass})";
                searcher.PropertiesToLoad.Add("distinguishedName");
                searcher.PropertiesToLoad.Add("cn");
                searcher.PropertiesToLoad.Add("name");
                searcher.PropertiesToLoad.Add("description");
                searcher.SearchScope = SearchScope.Subtree;

                foreach (SearchResult res in searcher.FindAll())
                {
                    list.Add(new AdObjectInfo
                    {
                        DistinguishedName = res.Properties["distinguishedName"]?.Count > 0 ? res.Properties["distinguishedName"][0].ToString() : null,
                        CommonName = res.Properties["cn"]?.Count > 0 ? res.Properties["cn"][0].ToString() : null,
                        Name = res.Properties["name"]?.Count > 0 ? res.Properties["name"][0].ToString() : null,
                        Description = res.Properties["description"]?.Count > 0 ? res.Properties["description"][0].ToString() : null,
                    });
                }
            }

            return list;
        }
    }
}
