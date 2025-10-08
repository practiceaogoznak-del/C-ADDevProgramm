using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
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

        public bool IsConnected()
        {
            try
            {
                using (var entry = new DirectoryEntry(_ldapPath))
                {
                    var _ = entry.NativeObject; // тест подключения
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        // ======== Получение ресурсов (группы / компьютеры) ========
        public List<AdObjectInfo> GetResources(string objectClass)
        {
            var list = new List<AdObjectInfo>();
            try
            {
                using (DirectoryEntry entry = new DirectoryEntry(_ldapPath))
                using (DirectorySearcher searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(objectClass={objectClass})";
                    searcher.PropertiesToLoad.Add("name");
                    searcher.PropertiesToLoad.Add("description");

                    foreach (SearchResult result in searcher.FindAll())
                    {
                        string name = result.Properties["name"].Count > 0 ? result.Properties["name"][0].ToString() : "";
                        string desc = result.Properties["description"].Count > 0 ? result.Properties["description"][0].ToString() : "";
                        list.Add(new AdObjectInfo { Name = name, Description = desc });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Ошибка при получении ресурсов: " + ex.Message);
            }

            return list;
        }

        // ======== Группы, где состоит пользователь ========
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
                        foreach (var g in user.GetGroups())
                            userGroups.Add(g.SamAccountName);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Ошибка при получении групп пользователя: " + ex.Message);
            }
            return userGroups;
        }

        // ======== Почта владельца (управляющего) группы ========
        public string GetGroupOwnerEmail(string groupName)
        {
            try
            {
                using (DirectoryEntry entry = new DirectoryEntry(_ldapPath))
                using (DirectorySearcher searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(&(objectClass=group)(sAMAccountName={groupName}))";
                    searcher.PropertiesToLoad.Add("managedBy");

                    var result = searcher.FindOne();
                    if (result != null && result.Properties["managedBy"].Count > 0)
                    {
                        string managedByDn = result.Properties["managedBy"][0].ToString();

                        // Берём именно владельца, не группу!
                        using (var ownerEntry = new DirectoryEntry("LDAP://" + managedByDn))
                        {
                            var mail = ownerEntry.Properties["mail"].Value?.ToString();
                            return mail;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Ошибка при получении владельца группы: " + ex.Message);
            }

            return null;
        }
    }

  
}
