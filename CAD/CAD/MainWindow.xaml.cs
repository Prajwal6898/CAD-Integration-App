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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CAD
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            
            // Enable window dragging
            this.MouseLeftButtonDown += MainWindow_MouseLeftButtonDown;
        }
        
        private void MainWindow_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Allow dragging the window by clicking anywhere on it
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }

        private void OpenPopupButton_Click(object sender, RoutedEventArgs e)
        {
            // Create and show the modern popup window
            PopupWindow popup = new PopupWindow();
            popup.Show(); // Show non-modal
            
            // Close the welcome window
            this.Close();
        }
    }
}
