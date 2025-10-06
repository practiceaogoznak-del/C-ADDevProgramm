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
        private readonly string _ldapPath;

        public ActiveDirectoryHelper(string ldapPath)
        {
            _ldapPath = ldapPath;
        }

        /// <summary>
        /// Получает список объектов указанного типа (group, computer, user и т.п.)
        /// </summary>
        public List<AdObjectInfo> GetResources(string objectCategory)
        {
            var results = new List<AdObjectInfo>();

            try
            {
                using (var entry = new DirectoryEntry(_ldapPath))
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(objectCategory={objectCategory})";
                    searcher.PropertiesToLoad.AddRange(new[]
                    {
                        "cn",
                        "name",
                        "description",
                        "displayName",
                        "mail",
                        "telephoneNumber",
                        "employeeID"
                    });
                    searcher.SizeLimit = 500;
                    searcher.PageSize = 500;

                    foreach (SearchResult result in searcher.FindAll())
                    {
                        var info = new AdObjectInfo
                        {
                            Name = GetProp(result, "name"),
                            DisplayName = GetProp(result, "displayName"),
                            Description = GetProp(result, "description"),
                            Email = GetProp(result, "mail"),
                            Telephone = GetProp(result, "telephoneNumber"),
                            EmployeeId = GetProp(result, "employeeID"),
                            DistinguishedName = GetProp(result, "distinguishedName")
                        };

                        results.Add(info);
                    }
                }
            }
            catch (Exception ex)
            {
                // Можно добавить логирование, чтобы не падало приложение
                System.Diagnostics.Debug.WriteLine("Ошибка AD: " + ex.Message);
            }

            return results;
        }

        private string GetProp(SearchResult result, string prop)
        {
            return result.Properties.Contains(prop)
                ? result.Properties[prop][0]?.ToString()
                : string.Empty;
        }
    }
}
