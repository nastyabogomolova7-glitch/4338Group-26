using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Group4338
{
    /// <summary>
    /// Логика взаимодействия для _4338_Bogomolova.xaml
    /// </summary>
public partial class _4338_Bogomolova : Window
        {
            public _4338_Bogomolova()
            {
                InitializeComponent();
            }

            private void CloseButton_Click(object sender, RoutedEventArgs e)
            {
                this.Close();
            }
        }
    }
