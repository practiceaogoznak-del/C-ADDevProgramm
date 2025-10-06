using System.Security.Principal;
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

namespace AccessManager
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {


        public MainWindow()
        {
            InitializeComponent();

            openForm();

        }
        public void openForm()
        {
            MainFrame.Navigate(new Views.RequestFormVIew());
        }

        public void Authorization()
        {
            string currentUser = WindowsIdentity.GetCurrent().Name;
            aut.Text = currentUser;
        }
    }
}