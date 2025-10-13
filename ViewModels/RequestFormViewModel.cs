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

        // Основная информация о пользователе
        public string ApplicantFullName { get; set; }
        public string ApplicantTabNumber { get; set; }
        public string Position { get; set; }
        public string ContactPhone { get; set; }

        // Поля формы
        public ObservableCollection<string> ActionTypes { get; set; }
        public string SelectedActionType { get; set; }
        public string Reason { get; set; }

        // Коллекции
        public ObservableCollection<ResourceRequestItem> SelectedRequestItems { get; set; }
        private ObservableCollection<ResourceRequestItem> _allRequestItems;
        private ObservableCollection<AdObjectInfo> _allResources;

        public ObservableCollection<AdObjectInfo> AvailableWorkstations { get; set; }
        public AdObjectInfo SelectedWorkstation { get; set; }

        // Временный доступ
        private bool _isTemporary;
        public bool IsTemporary
        {
            get => _isTemporary;
            set
            {
                if (_isTemporary != value)
                {
                    _isTemporary = value;
                    OnPropertyChanged(nameof(IsTemporary));
                }
            }
        }
        public DateTime? TemporaryFrom { get; set; }
        public DateTime? TemporaryUntil { get; set; }

        // Фильтры
        private bool _showOnlyWithEmail;
        public bool ShowOnlyWithEmail
        {
            get => _showOnlyWithEmail;
            set
            {
                _showOnlyWithEmail = value;
                OnPropertyChanged(nameof(ShowOnlyWithEmail));
                FilterResources();
            }
        }

        private bool _showOnlyAccessible;
        public bool ShowOnlyAccessible
        {
            get => _showOnlyAccessible;
            set
            {
                _showOnlyAccessible = value;
                OnPropertyChanged(nameof(ShowOnlyAccessible));
                FilterResources();
            }
        }

        // Поиск
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

        // Команды
        public ICommand SaveDraftCommand { get; }
        public ICommand SubmitCommand { get; }
        public ICommand SearchResourcesCommand { get; }
        public ICommand FilterByEmailCommand { get; }
        public ICommand FilterByAccessCommand { get; }

        // Конструктор
        public RequestFormViewModel()
        {
         

            string ldapPath = GetDefaultLdapPath();
            _adHelper = new ActiveDirectoryHelper(ldapPath);
            _emailService = new EmailService();

            ActionTypes = new ObservableCollection<string> { "Добавить", "Удалить", "Изменить" };
            SelectedActionType = ActionTypes.FirstOrDefault();

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
            FilterByEmailCommand = new RelayCommand(_ => FilterResources());
            FilterByAccessCommand = new RelayCommand(_ => FilterResources());
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

                // Определяем владельцев
                foreach (var resource in _allResources)
                {
                    try
                    {
                        string email = _adHelper.GetGroupOwnerEmail(resource.Name);
                        if (!string.IsNullOrEmpty(email))
                        {
                            resource.Email = email;
                            resource.HasOwner = true;
                        }
                        else
                        {
                            resource.Email = null;
                            resource.HasOwner = false;
                        }
                    }
                    catch (Exception ex)
                    {
                        resource.HasOwner = false;
                        MessageBox.Show($"Ошибка при получении владельца для {resource.Name}: {ex.Message}");
                    }
                }

                string userName = Environment.UserName;
                var userGroups = _adHelper.GetUserGroups(userName)
                    .Select(g => g.ToLowerInvariant())
                    .ToList();

                _allRequestItems = new ObservableCollection<ResourceRequestItem>(
                    _allResources.Select(r => new ResourceRequestItem
                    {
                        Resource = r,
                        CurrentHasAccess = userGroups.Contains(r.Name.ToLowerInvariant()),
                        AccessLevel = userGroups.Contains(r.Name.ToLowerInvariant()) ? "Чтение и запись" : "Нет доступа",
                        IsRequested = false
                    })
                );

                SelectedRequestItems = new ObservableCollection<ResourceRequestItem>(_allRequestItems);

                var computers = _adHelper.GetResources("computer");
                AvailableWorkstations = new ObservableCollection<AdObjectInfo>(computers);

                string computerName = Environment.MachineName;
                SelectedWorkstation = AvailableWorkstations
                    .FirstOrDefault(w => w.Name.Equals(computerName, StringComparison.OrdinalIgnoreCase));

               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке AD: " + ex.Message);
            }
        }

        private void SearchResources()
        {
            try
            {
                IEnumerable<ResourceRequestItem> filtered;

                if (string.IsNullOrWhiteSpace(ResourceSearchText))
                    filtered = _allRequestItems;
                else
                    filtered = _allRequestItems.Where(r =>
                        r.Resource.Name.Contains(ResourceSearchText, StringComparison.OrdinalIgnoreCase) ||
                        (r.Resource.Description?.Contains(ResourceSearchText, StringComparison.OrdinalIgnoreCase) ?? false));

                SelectedRequestItems = new ObservableCollection<ResourceRequestItem>(filtered);
                OnPropertyChanged(nameof(SelectedRequestItems));

               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при поиске ресурсов: " + ex.Message);
            }
        }

        private void FilterResources()
        {
            try
            {
                var filtered = _allRequestItems.AsEnumerable();

                if (ShowOnlyWithEmail)
                    filtered = filtered.Where(r => !string.IsNullOrEmpty(r.Resource.Email));

                if (ShowOnlyAccessible)
                    filtered = filtered.Where(r => r.CurrentHasAccess);

                SelectedRequestItems = new ObservableCollection<ResourceRequestItem>(filtered);
                OnPropertyChanged(nameof(SelectedRequestItems));

             
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при фильтрации: " + ex.Message);
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
                    $"Причина: {Reason}\n";

                if (IsTemporary)
                {
                    body += $"Прошу предоставить доступ с {TemporaryFrom?.ToShortDateString()} по {TemporaryUntil?.ToShortDateString()}\n";
                }

                body += "\nВыбранные ресурсы:\n" + string.Join("\n", selectedResources);

                string to = string.Join(";", ownersEmails.Distinct());
                _emailService.CreateOutlookEmail(to, subject, body);

             
                MessageBox.Show($"Письмо успешно сформировано для владельцев:\n{to}");
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

    public class ResourceRequestItem
    {
        public AdObjectInfo Resource { get; set; }
        public bool CurrentHasAccess { get; set; }
        public string AccessLevel { get; set; }
        public bool IsRequested { get; set; }
    }
}
