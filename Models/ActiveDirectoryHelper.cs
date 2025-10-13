using AccessManager.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Windows;

namespace AccessManager.Models
{
    public class SyncResult
    {
        public List<UserProfileModel> ADUsers { get; set; } = new List<UserProfileModel>();
        public List<UserProfileEntity> NewlyInserted { get; set; } = new List<UserProfileEntity>();
        public List<UserProfileEntity> Updated { get; set; } = new List<UserProfileEntity>();
        public List<UserProfileEntity> AllSaved { get; set; } = new List<UserProfileEntity>();
    }

    public static class ActiveDirectoryHelper
    {
        public static SyncResult GetListUsersAD()
        {
            string container1 = "OU=Пермская печатная фабрика,OU=PPF,DC=gz,DC=local";
            string container2 = "OU=Dead Records,OU=PPF,DC=gz,DC=local";

            var list = GetUsersFromContainer(container1);
            var disabled = GetUsersFromContainer(container2);

            if (disabled.Count > 0)
                list.AddRange(disabled);

            var result = SaveUsersToDatabase(list);
            result.ADUsers = list;
            return result;
        }

        private static List<UserProfileModel> GetUsersFromContainer(string container)
        {
            var result = new List<UserProfileModel>();
            string domain = "06dc03.gz.local";

            using (var context = new PrincipalContext(ContextType.Domain, domain, container))
            using (var userPrincipal = new UserPrincipal(context))
            using (var search = new PrincipalSearcher(userPrincipal))
            {
                foreach (var principal in search.FindAll().OfType<UserPrincipal>())
                {
                    var entry = principal.GetUnderlyingObject() as DirectoryEntry;
                    if (entry != null)
                    {
                        result.Add(new UserProfileModel
                        {
                            Login = principal.SamAccountName ?? "—",
                            FullName = principal.DisplayName ?? "—",
                            Description = entry.Properties["description"].Value?.ToString() ?? "",
                            Enabled = principal.Enabled ?? false
                        });
                    }
                }

            }
            return result;
        }

        public static SyncResult SaveUsersToDatabase(List<UserProfileModel> users)
        {
            var result = new SyncResult();

            try
            {
                using (var db = new UserDbContext())
                {
                    // Подготовим словарь существующих пользователей для одной выборки (повышаем производительность)
                    var logins = users.Select(u => u.Login).Where(l => !string.IsNullOrEmpty(l)).Distinct().ToList();
                    var existingDict = db.UsersAD.Where(u => logins.Contains(u.Login)).ToDictionary(u => u.Login, u => u);

                    foreach (var user in users)
                    {
                        if (string.IsNullOrEmpty(user.Login))
                            continue;

                        if (existingDict.TryGetValue(user.Login, out var existing))
                        {
                            bool changed = existing.FullName != user.FullName
                                           || existing.Description != user.Description
                                           || existing.Enabled != user.Enabled;

                            existing.FullName = user.FullName;
                            existing.Description = user.Description;
                            existing.Enabled = user.Enabled;
                            existing.UpdatedAt = DateTime.Now;

                            if (changed)
                                result.Updated.Add(existing);
                        }
                        else
                        {
                            var entity = new UserProfileEntity
                            {
                                Login = user.Login,
                                FullName = user.FullName,
                                Description = user.Description,
                                Enabled = user.Enabled,
                                UpdatedAt = DateTime.Now
                            };
                            db.UsersAD.Add(entity);
                            result.NewlyInserted.Add(entity);
                        }
                    }

                    db.SaveChanges();

                    // Получим актуальные записи из БД по списку логинов
                    result.AllSaved = db.UsersAD.AsNoTracking().Where(u => logins.Contains(u.Login)).ToList();
                }
            }
            catch (Exception ex)
            {
                var errorMessage = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
                MessageBox.Show($"Ошибка при сохранении: {errorMessage}", "Ошибка");
            }

            return result;
        }
    }
}
