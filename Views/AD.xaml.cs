
using AccessManager.Models;
using System;
using System.ComponentModel;
using System.Data.Entity;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AccessManager.Views
{
    /// <summary>
    /// Логика взаимодействия для AD.xaml
    /// </summary>
    public partial class AD : Page
    {
        private List<UserProfileModel> _adUsers = new List<UserProfileModel>();
        private List<UserProfileEntity> _savedUsers = new List<UserProfileEntity>();

        private ICollectionView _viewAd;
        private ICollectionView _viewSaved;

        public AD()
        {
            InitializeComponent();
            cmbSort.SelectedIndex = 0;
        }

        private void btnLoad_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Выполнить обновление списка пользователей из AD?",
                    "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                    return;

                // Выполняем синхронизацию и получаем результат
                var sync = ActiveDirectoryHelper.GetListUsersAD();

                _adUsers = sync.ADUsers ?? new List<UserProfileModel>();
                _savedUsers = sync.NewlyInserted ?? new List<UserProfileEntity>();

                // Привязываем списки к ListBox и настраиваем view для фильтра и сортировки
                listAD.ItemsSource = _adUsers;
                _viewAd = CollectionViewSource.GetDefaultView(listAD.ItemsSource);
                _viewAd.Filter = FilterBySearch;

                listSaved.ItemsSource = _savedUsers;
                _viewSaved = CollectionViewSource.GetDefaultView(listSaved.ItemsSource);
                _viewSaved.Filter = FilterBySearchForSaved;

                ApplySort();

                // Обновляем отображение
                _viewAd?.Refresh();
                _viewSaved?.Refresh();

                MessageBox.Show($"Пользователей в AD: {_adUsers.Count}\nНовых записано в БД: {_savedUsers.Count}", "Готово");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
            }
        }

        private void txtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            _viewAd?.Refresh();
            _viewSaved?.Refresh();
        }

        private void btnClearSearch_Click(object sender, RoutedEventArgs e)
        {
            txtSearch.Text = string.Empty;
            _viewAd?.Refresh();
            _viewSaved?.Refresh();
        }

        private void cmbSort_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplySort();
        }

        private void ApplySort()
        {
            if (_viewAd == null || _viewSaved == null) return;

            // Сброс сортировок
            _viewAd.SortDescriptions.Clear();
            _viewSaved.SortDescriptions.Clear();

            switch (cmbSort.SelectedIndex)
            {
                case 0: // по фамилии А→Я
                    _viewAd.SortDescriptions.Add(new SortDescription(nameof(UserProfileModel.FullName), ListSortDirection.Ascending));
                    _viewSaved.SortDescriptions.Add(new SortDescription(nameof(UserProfileEntity.FullName), ListSortDirection.Ascending));
                    break;
                case 1: // по фамилии Я→А
                    _viewAd.SortDescriptions.Add(new SortDescription(nameof(UserProfileModel.FullName), ListSortDirection.Descending));
                    _viewSaved.SortDescriptions.Add(new SortDescription(nameof(UserProfileEntity.FullName), ListSortDirection.Descending));
                    break;
                case 2: // по логину
                    _viewAd.SortDescriptions.Add(new SortDescription(nameof(UserProfileModel.Login), ListSortDirection.Ascending));
                    _viewSaved.SortDescriptions.Add(new SortDescription(nameof(UserProfileEntity.Login), ListSortDirection.Ascending));
                    break;
            }
        }

        // Фильтр по фамилии для AD (проверяет фамилию (первое слово в FullName) и FullName целиком)
        private bool FilterBySearch(object item)
        {
            if (string.IsNullOrWhiteSpace(txtSearch.Text)) return true;
            var s = txtSearch.Text.Trim();

            if (item is UserProfileModel u)
            {
                var fullname = u.FullName ?? string.Empty;
                var surname = fullname.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty;

                return surname.StartsWith(s, StringComparison.OrdinalIgnoreCase)
                    || (!string.IsNullOrEmpty(fullname) && fullname.IndexOf(s, StringComparison.OrdinalIgnoreCase) >= 0)
                    || (!string.IsNullOrEmpty(u.Login) && u.Login.IndexOf(s, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            return true;
        }

        // Фильтр по фамилии для списка только что записанных
        private bool FilterBySearchForSaved(object item)
        {
            if (string.IsNullOrWhiteSpace(txtSearch.Text)) return true;
            var s = txtSearch.Text.Trim();

            if (item is UserProfileEntity u)
            {
                var fullname = u.FullName ?? string.Empty;
                var surname = fullname.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty;

                return surname.StartsWith(s, StringComparison.OrdinalIgnoreCase)
                    || (!string.IsNullOrEmpty(fullname) && fullname.IndexOf(s, StringComparison.OrdinalIgnoreCase) >= 0)
                    || (!string.IsNullOrEmpty(u.Login) && u.Login.IndexOf(s, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            return true;
        }

        private void btnNext_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}

