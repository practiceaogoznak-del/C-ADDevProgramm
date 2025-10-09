using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using AccessManager.Classes;
using AccessManager.Services;

namespace AccessManager.ViewModels
{
    public class RequestFormViewModel : INotifyPropertyChanged
    {
        private readonly EmailService _emailService;
        private readonly ActiveDirectoryHelper _adHelper;
        private readonly string _logFilePath;

        public string ApplicantFullName { get; set; }
        public string ApplicantTabNumber { get; set; }
        public string Position { get; set; }
        public string ContactPhone { get; set; }

        public ObservableCollection<string> ActionTypes { get; set; }
        public string SelectedActionType { get; set; }
        public string Reason { get; set; }

        public ObservableCollection<ResourceRequestItem> SelectedRequestItems { get; set; }
        public ObservableCollection<AdObjectInfo> ResourceCatalog { get; set; }
        public ObservableCollection<AdObjectInfo> AvailableWorkstations { get; set; }

        private ObservableCollection<AdObjectInfo> _allResources;
        public AdObjectInfo SelectedWorkstation { get; set; }

        public bool IsTemporary { get; set; }
        public DateTime? TemporaryUntil { get; set; }

        private string _resourceSearchText;
        public string ResourceSearchText
        {
            get => _resourceSearchText;
            set
            {
                _resourceSearchText = value;
                OnPropertyChanged(nameof(ResourceSearchText));
            }
        }

        public ICommand SaveDraftCommand { get; }
        public ICommand SubmitCommand { get; }
        public ICommand SearchResourcesCommand { get; }

        public RequestFormViewModel()
        {
            _logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AccessManager.log");

            string ldapPath = GetDefaultLdapPath();
            _adHelper = new ActiveDirectoryHelper(ldapPath, _logFilePath);
            _emailService = new EmailService();

            ActionTypes = new ObservableCollection<string> { "Добавить", "Удалить", "Изменить" };
            SelectedActionType = ActionTypes.FirstOrDefault();

            ResourceCatalog = new ObservableCollection<AdObjectInfo>();
            SelectedRequestItems = new ObservableCollection<ResourceRequestItem>();
            AvailableWorkstations = new ObservableCollection<AdObjectInfo>();

            if (_adHelper.IsConnected())
            {
                LoadUserInfoFromAD();
                LoadAdData();
            }
            else
            {
                MessageBox.Show("Не удалось подключиться к AD.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                ApplicantFullName = Environment.UserName;
            }

            SaveDraftCommand = new RelayCommand(_ => SaveDraft());
            SubmitCommand = new RelayCommand(_ => Submit());
            SearchResourcesCommand = new RelayCommand(_ => SearchResources());
        }

        private string GetDefaultLdapPath()
        {
            try
            {
                DirectoryEntry rootDse = new DirectoryEntry("LDAP://RootDSE");
                string defaultNamingContext = (string)rootDse.Properties["defaultNamingContext"].Value;
                return "LDAP://" + defaultNamingContext;
            }
            catch
            {
                return "LDAP://";
            }
        }

        private void LoadAdData()
        {
            try
            {
                Log("Загрузка данных из AD...");
                var groups = _adHelper.GetResources("group");
                _allResources = new ObservableCollection<AdObjectInfo>(groups);
                ResourceCatalog = new ObservableCollection<AdObjectInfo>(_allResources);

                string userName = Environment.UserName;
                var userGroups = _adHelper.GetUserGroups(userName)
                    .Select(g => g.ToLowerInvariant())
                    .ToList();

                SelectedRequestItems = new ObservableCollection<ResourceRequestItem>(
                    _allResources.Select(r => new ResourceRequestItem
                    {
                        Resource = r,
                        CurrentHasAccess = userGroups.Contains(r.Name.ToLowerInvariant()),
                        IsRequested = userGroups.Contains(r.Name.ToLowerInvariant())
                    })
                );

                var computers = _adHelper.GetResources("computer");
                AvailableWorkstations = new ObservableCollection<AdObjectInfo>(computers);

                string computerName = Environment.MachineName;
                SelectedWorkstation = AvailableWorkstations
                    .FirstOrDefault(w => w.Name.Equals(computerName, StringComparison.OrdinalIgnoreCase));

                Log("Загрузка AD завершена.");
            }
            catch (Exception ex)
            {
                Log("Ошибка при загрузке AD: " + ex.Message);
            }
        }

        private void SearchResources()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ResourceSearchText))
                {
                    ResourceCatalog = new ObservableCollection<AdObjectInfo>(_allResources);
                }
                else
                {
                    var filtered = _allResources
                        .Where(r => r.Name.Contains(ResourceSearchText, StringComparison.OrdinalIgnoreCase) ||
                                    (r.Description?.Contains(ResourceSearchText, StringComparison.OrdinalIgnoreCase) ?? false))
                        .ToList();

                    ResourceCatalog = new ObservableCollection<AdObjectInfo>(filtered);
                }

                OnPropertyChanged(nameof(ResourceCatalog));
                Log($"Поиск ресурсов по запросу: \"{ResourceSearchText}\" — найдено {ResourceCatalog.Count}");
            }
            catch (Exception ex)
            {
                Log("Ошибка при поиске ресурсов: " + ex.Message);
            }
        }

        private void LoadUserInfoFromAD()
        {
            try
            {
                using (var context = new PrincipalContext(ContextType.Domain))
                using (var user = UserPrincipal.Current)
                {
                    ApplicantFullName = user?.DisplayName ?? Environment.UserName;
                    Position = user?.Description ?? "Не указано";
                    ContactPhone = user?.VoiceTelephoneNumber ?? "—";
                    ApplicantTabNumber = user?.EmployeeId ?? "—";
                }
            }
            catch (Exception ex)
            {
                Log("Ошибка при загрузке данных пользователя: " + ex.Message);
                ApplicantFullName = Environment.UserName;
            }
        }

        private void SaveDraft()
        {
            MessageBox.Show("Черновик сохранён (демо).");
        }

        private void Submit()
        {
            try
            {
                var selectedResources = SelectedRequestItems
                    .Where(i => i.IsRequested)
                    .Select(i => i.Resource.Name)
                    .ToList();

                if (selectedResources.Count == 0)
                {
                    MessageBox.Show("Выберите хотя бы один ресурс.");
                    return;
                }

                var ownersEmails = new List<string>();
                foreach (var resource in selectedResources)
                {
                    string ownerMail = _adHelper.GetGroupOwnerEmail(resource);
                    if (!string.IsNullOrEmpty(ownerMail))
                        ownersEmails.Add(ownerMail);
                }

                if (ownersEmails.Count == 0)
                {
                    MessageBox.Show("Не удалось определить владельцев выбранных ресурсов.");
                    return;
                }

                string subject = "Заявка на доступ";
                string body =
                    $"ФИО: {ApplicantFullName}\n" +
                    $"Должность: {Position}\n" +
                    $"Табельный №: {ApplicantTabNumber}\n" +
                    $"Телефон: {ContactPhone}\n" +
                    $"Действие: {SelectedActionType}\n" +
                    $"Причина: {Reason}\n" +
                    $"Временно: {(IsTemporary ? "Да" : "Нет")}\n" +
                    (IsTemporary ? $"Дата окончания: {TemporaryUntil?.ToShortDateString()}\n" : "") +
                    "\nВыбранные ресурсы:\n" +
                    string.Join("\n", selectedResources);

                string to = string.Join(";", ownersEmails.Distinct());
                _emailService.CreateOutlookEmail(to, subject, body);

                Log($"Формирование письма успешно: получатели — {to}");
                MessageBox.Show($"Письмо успешно сформировано для владельцев:\n{to}");
            }
            catch (Exception ex)
            {
                Log("Ошибка при отправке письма: " + ex.Message);
                MessageBox.Show("Ошибка при отправке: " + ex.Message);
            }
        }

        private void Log(string message)
        {
            File.AppendAllText(_logFilePath, $"[{DateTime.Now}] {message}\n");
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
