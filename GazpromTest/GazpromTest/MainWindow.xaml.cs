using GazpromTest.Models;
using System.Windows;

namespace GazpromTest
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            ServiceManager.Init();
            InitializeComponent();
        }
    }
}
