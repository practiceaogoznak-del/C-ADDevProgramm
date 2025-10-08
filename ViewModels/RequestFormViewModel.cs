using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
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

        // ================= Основные поля =================
        public string ApplicantFullName { get; set; }
        public string ApplicantTabNumber { get; set; }
        public string Position { get; set; }
        public string ContactPhone { get; set; }

        public ObservableCollection<string> ActionTypes { get; set; }
        public string SelectedActionType { get; set; }
        public string Reason { get; set; }

        // ================= Коллекции =================
        public ObservableCollection<ResourceRequestItem> SelectedRequestItems { get; set; }
        public ObservableCollection<AdObjectInfo> ResourceCatalog { get; set; }
        public ObservableCollection<AdObjectInfo> AvailableWorkstations { get; set; }

        private ObservableCollection<AdObjectInfo> _allResources;

        public AdObjectInfo SelectedWorkstation { get; set; }

        // ================= Флаги =================
        public bool IsTemporary { get; set; }
        public DateTime? TemporaryUntil { get; set; }

        // ================= Команды =================
        public ICommand SaveDraftCommand { get; }
        public ICommand SubmitCommand { get; }

        public RequestFormViewModel()
        {
            string ldapPath = GetDefaultLdapPath();
            _adHelper = new ActiveDirectoryHelper(ldapPath);
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
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Ошибка при загрузке AD: " + ex.Message);
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
            catch
            {
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
                _emailService.SendOutlookEmail(to, subject, body);

                MessageBox.Show($"Письмо успешно отправлено владельцам:\n{to}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при отправке: " + ex.Message);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // ===== Класс элемента ресурса =====
    public class ResourceRequestItem : INotifyPropertyChanged
    {
        public AdObjectInfo Resource { get; set; }

        private bool _isRequested;
        public bool IsRequested
        {
            get => _isRequested;
            set { _isRequested = value; OnPropertyChanged(nameof(IsRequested)); }
        }

        private bool _currentHasAccess;
        public bool CurrentHasAccess
        {
            get => _currentHasAccess;
            set { _currentHasAccess = value; OnPropertyChanged(nameof(CurrentHasAccess)); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        }
    }
}
