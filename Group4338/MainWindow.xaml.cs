using System.Windows;

namespace Group4338
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void AuthorButton_Click(object sender, RoutedEventArgs e)
        {
            var authorWindow = new _4338_Bogomolova();
            authorWindow.ShowDialog();
        }
    }
}