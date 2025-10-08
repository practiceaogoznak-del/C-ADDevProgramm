using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Windows.Input;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using AccessManager.Services;
using AccessManager.Classes;
using System.Windows;

namespace AccessManager.ViewModels
{
    public class RequestFormViewModel : INotifyPropertyChanged
    {
        private readonly ActiveDirectoryHelper _adHelper;

        // ================= Основные поля заявки =================

        private string _applicantFullName;
        public string ApplicantFullName
        {
            get => _applicantFullName;
            set { _applicantFullName = value; OnPropertyChanged(nameof(ApplicantFullName)); }
        }

        private string _applicantTabNumber;
        public string ApplicantTabNumber
        {
            get => _applicantTabNumber;
            set { _applicantTabNumber = value; OnPropertyChanged(nameof(ApplicantTabNumber)); }
        }

        private string _position;
        public string Position
        {
            get => _position;
            set { _position = value; OnPropertyChanged(nameof(Position)); }
        }

        private string _contactPhone;
        public string ContactPhone
        {
            get => _contactPhone;
            set { _contactPhone = value; OnPropertyChanged(nameof(ContactPhone)); }
        }

        public ObservableCollection<string> ActionTypes { get; set; }
        private string _selectedActionType;
        public string SelectedActionType
        {
            get => _selectedActionType;
            set { _selectedActionType = value; OnPropertyChanged(nameof(SelectedActionType)); }
        }

        private string _reason;
        public string Reason
        {
            get => _reason;
            set { _reason = value; OnPropertyChanged(nameof(Reason)); }
        }

        // ================= Ресурсы =================

        private ObservableCollection<ResourceRequestItem> _selectedRequestItems;
        public ObservableCollection<ResourceRequestItem> SelectedRequestItems
        {
            get => _selectedRequestItems;
            set { _selectedRequestItems = value; OnPropertyChanged(nameof(SelectedRequestItems)); }
        }

        private ObservableCollection<AdObjectInfo> _resourceCatalog;
        public ObservableCollection<AdObjectInfo> ResourceCatalog
        {
            get => _resourceCatalog;
            set { _resourceCatalog = value; OnPropertyChanged(nameof(ResourceCatalog)); }
        }

        private AdObjectInfo _selectedResource;
        public AdObjectInfo SelectedResource
        {
            get => _selectedResource;
            set { _selectedResource = value; OnPropertyChanged(nameof(SelectedResource)); }
        }

        // ================= Поиск по ресурсам =================
        private string _resourceSearchText;
        public string ResourceSearchText
        {
            get => _resourceSearchText;
            set
            {
                _resourceSearchText = value;
                OnPropertyChanged(nameof(ResourceSearchText));
                ApplyResourceFilter();
            }
        }

        private ObservableCollection<AdObjectInfo> _allResources;

        // ================= Рабочие места =================
        private ObservableCollection<AdObjectInfo> _availableWorkstations;
        public ObservableCollection<AdObjectInfo> AvailableWorkstations
        {
            get => _availableWorkstations;
            set { _availableWorkstations = value; OnPropertyChanged(nameof(AvailableWorkstations)); }
        }

        private AdObjectInfo _selectedWorkstation;
        public AdObjectInfo SelectedWorkstation
        {
            get => _selectedWorkstation;
            set { _selectedWorkstation = value; OnPropertyChanged(nameof(SelectedWorkstation)); }
        }

        // ================= Временное / постоянное =================
        private bool _isTemporary;
        public bool IsTemporary
        {
            get => _isTemporary;
            set { _isTemporary = value; OnPropertyChanged(nameof(IsTemporary)); }
        }

        private DateTime? _temporaryUntil;
        public DateTime? TemporaryUntil
        {
            get => _temporaryUntil;
            set { _temporaryUntil = value; OnPropertyChanged(nameof(TemporaryUntil)); }
        }

        // ================= Команды =================
        public ICommand SaveDraftCommand { get; set; }
        public ICommand SubmitCommand { get; set; }

        // ================= Конструктор =================
        public RequestFormViewModel()
        {
            string ldapPath = GetDefaultLdapPath();
            _adHelper = new ActiveDirectoryHelper(ldapPath);

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
                MessageBox.Show(
                    "Не удалось подключиться к Active Directory.\n" +
                    "Проверьте сетевое соединение или подключение к домену.",
                    "Ошибка LDAP подключения",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );

                // fallback-заполнение локальными данными
                ApplicantFullName = Environment.UserName;
                Position = "Локальный пользователь";
                ContactPhone = "—";
                ApplicantTabNumber = "—";
            }

            SaveDraftCommand = new RelayCommand(_ => SaveDraft(), _ => CanSaveDraft());
            SubmitCommand = new RelayCommand(_ => Submit(), _ => CanSubmit());
        }

        // ================= Получение LDAP пути =================
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

        // ================= Загрузка данных AD =================
        private void LoadAdData()
        {
            try
            {
                // Загружаем все ресурсы
                var groups = _adHelper.GetResources("group");
                _allResources = new ObservableCollection<AdObjectInfo>(groups);
                ResourceCatalog = new ObservableCollection<AdObjectInfo>(_allResources);

                // Получаем текущего пользователя
                string userName = Environment.UserName; // или можно user.SamAccountName
                var userGroups = _adHelper.GetUserGroups(userName)
                    .Select(g => g.ToLowerInvariant())
                    .ToList();

                // Формируем список ресурсов с флагом доступа
                SelectedRequestItems = new ObservableCollection<ResourceRequestItem>(
                    _allResources.Select(r => new ResourceRequestItem
                    {
                        Resource = r,
                        CurrentHasAccess = userGroups.Contains(r.Name.ToLowerInvariant()),
                        IsRequested = userGroups.Contains(r.Name.ToLowerInvariant()) // чтобы сразу отметить чекбокс
                    })
                );

                // Загрузка компьютеров
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


        // ================= Поиск по ресурсам =================
        private void ApplyResourceFilter()
        {
            if (_allResources == null) return;

            if (string.IsNullOrWhiteSpace(ResourceSearchText))
            {
                ResourceCatalog = new ObservableCollection<AdObjectInfo>(_allResources);
            }
            else
            {
                var filtered = _allResources
                    .Where(r =>
                        (r.Name != null && r.Name.Contains(ResourceSearchText, StringComparison.OrdinalIgnoreCase)) ||
                        (r.Description != null && r.Description.Contains(ResourceSearchText, StringComparison.OrdinalIgnoreCase)))
                    .ToList();

                ResourceCatalog = new ObservableCollection<AdObjectInfo>(filtered);
            }

            OnPropertyChanged(nameof(ResourceCatalog));
        }

        // ================= Загрузка пользователя из AD =================
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
                // fallback если AD недоступен
                ApplicantFullName = Environment.UserName;
                Position = "Локальный пользователь";
                ContactPhone = "—";
                ApplicantTabNumber = "—";
            }
        }

        // ================= Черновик / Отправка =================
        private void SaveDraft()
        {
            // Сохраняешь черновик — например, в файл
        }

        private bool CanSaveDraft() => true;

        private void Submit()
        {
            // Реализация отправки заявки
        }

        private bool CanSubmit() => true;

        // ================= PropertyChanged =================
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        }
    }

    // ================= Класс элемента ресурса =================
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
