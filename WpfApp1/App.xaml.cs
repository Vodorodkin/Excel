using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для App.xaml
    /// </summary>
    public partial class App : Application
    {

        protected override void OnStartup(StartupEventArgs e)
        {
            new MainWindow()
            {
                DataContext = new VMmain(),
                //Width = 300,
                //Height = 300,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Title = "Авторизация",
            }.Show();
        }


    }
}
