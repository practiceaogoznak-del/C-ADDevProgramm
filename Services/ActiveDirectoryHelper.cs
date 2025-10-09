using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using AccessManager.Classes;

namespace AccessManager.Services
{
    public class ActiveDirectoryHelper
    {


        private readonly string _ldapPath;
        private readonly string _logFilePath;

        public ActiveDirectoryHelper(string ldapPath, string logFilePath = null)
        {
            _ldapPath = ldapPath;
            _logFilePath = logFilePath;
        }

        private void Log(string text)
        {
            if (string.IsNullOrEmpty(_logFilePath)) return;
            File.AppendAllText(_logFilePath, $"[{DateTime.Now}] [AD] {text}\n");
        }

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
                    if (result == null)
                    {
                        Log($"Группа '{groupName}' не найдена.");
                        return null;
                    }

                    if (result.Properties["managedBy"].Count == 0)
                    {
                        Log($"У группы '{groupName}' не указано managedBy.");
                        return null;
                    }

                    string managedByDn = result.Properties["managedBy"][0].ToString();
                    return ResolveOwnerEmail(managedByDn, 0);
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка при поиске владельца {groupName}: {ex}");
                return null;
            }
        }

        private string ResolveOwnerEmail(string dn, int depth)
        {
            if (depth > 5) return null;
            try
            {
                using (var entry = new DirectoryEntry("LDAP://" + dn))
                {
                    string mail = entry.Properties["mail"].Value?.ToString();
                    string type = entry.SchemaClassName;
                    Log($"Объект '{entry.Properties["name"].Value}' ({type}), mail={mail ?? "нет"}");

                    if (!string.IsNullOrEmpty(mail))
                        return mail;

                    if (type.Equals("group", StringComparison.OrdinalIgnoreCase))
                    {
                        var managedBy = entry.Properties["managedBy"].Value?.ToString();
                        if (!string.IsNullOrEmpty(managedBy))
                            return ResolveOwnerEmail(managedBy, depth + 1);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка при обработке DN '{dn}': {ex.Message}");
            }
            return null;
        }

        public List<AdObjectInfo> GetResources(string objectClass)
        {
            var result = new List<AdObjectInfo>();

            try
            {
                using (var entry = new DirectoryEntry(_ldapPath))
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(objectClass={objectClass})";
                    searcher.PropertiesToLoad.Add("name");
                    searcher.PropertiesToLoad.Add("description");

                    foreach (SearchResult sr in searcher.FindAll())
                    {
                        string name = sr.Properties["name"].Count > 0 ? sr.Properties["name"][0].ToString() : "";
                        string desc = sr.Properties["description"].Count > 0 ? sr.Properties["description"][0].ToString() : "";

                        result.Add(new AdObjectInfo
                        {
                            Name = name,
                            Description = desc
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Ошибка при получении ресурсов из AD: " + ex.Message);
            }

            return result;
        }

        public List<string> GetUserGroups(string userSamAccountName)
        {
            var groups = new List<string>();

            try
            {
                using (var entry = new DirectoryEntry(_ldapPath))
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(&(objectClass=user)(sAMAccountName={userSamAccountName}))";
                    searcher.PropertiesToLoad.Add("memberOf");

                    var result = searcher.FindOne();
                    if (result != null && result.Properties["memberOf"] != null)
                    {
                        foreach (var dn in result.Properties["memberOf"])
                        {
                            string groupName = dn.ToString();
                            int start = groupName.IndexOf("CN=") + 3;
                            int end = groupName.IndexOf(",", start);
                            if (start >= 0 && end > start)
                            {
                                groupName = groupName.Substring(start, end - start);
                                groups.Add(groupName);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Ошибка при получении групп пользователя: " + ex.Message);
            }

            return groups;
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
    }


}
