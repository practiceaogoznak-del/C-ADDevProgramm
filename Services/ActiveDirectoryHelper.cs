using System;
using System.Collections.Generic;
using System.DirectoryServices;
using AccessManager.Classes;
using System.Windows;
using System.DirectoryServices.AccountManagement;

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
        /// Проверяет доступность LDAP соединения
        /// </summary>
        public bool IsConnected()
        {
            try
            {
                using (var entry = new DirectoryEntry(_ldapPath))
                {
                    var native = entry.NativeObject; // Пытаемся получить COM-объект
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("LDAP недоступен: " + ex.Message);
                return false;
            }
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
                        "employeeID",
                        "distinguishedName"
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

        public List<string> GetUserGroups(string userSamAccountName)
        {
            var userGroups = new List<string>();
            try
            {
                using (var context = new PrincipalContext(ContextType.Domain))
                using (var user = UserPrincipal.FindByIdentity(context, userSamAccountName))
                {
                    if (user != null)
                    {
                        var groups = user.GetGroups();
                        foreach (var g in groups)
                        {
                            userGroups.Add(g.SamAccountName); // или g.Name
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Ошибка при получении групп пользователя: " + ex.Message);
            }

            return userGroups;
        }
    }
}
