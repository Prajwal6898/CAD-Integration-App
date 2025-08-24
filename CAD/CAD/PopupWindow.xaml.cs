using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;

namespace CAD
{
    /// <summary>
    /// Interaction logic for PopupWindow.xaml
    /// </summary>
    public partial class PopupWindow : Window
    {
        private AutoCADConnection acadConnection;

        public PopupWindow()
        {
            InitializeComponent();
            InitializeAutoCADConnection();
        }

        private void InitializeAutoCADConnection()
        {
            try
            {
                acadConnection = new AutoCADConnection();
            }
            catch (Exception ex)
            {
                // Handle initialization error gracefully
                MessageBox.Show($"Warning: AutoCAD/ZWCAD connection initialization failed: {ex.Message}\n\nThe application will continue to work, but AutoCAD integration features may not be available.", 
                    "AutoCAD Connection Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                acadConnection = null;
            }
        }





        private void Close_Click(object sender, RoutedEventArgs e)
        {
            // Close the popup window
            this.Close();
        }

        #region AutoCAD/ZWCAD Integration Event Handlers

        private void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            if (acadConnection == null)
            {
                MessageBox.Show("AutoCAD connection is not available. Please restart the application.", 
                                "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                if (acadConnection.Connect())
                {
                    UpdateConnectionStatus(true, acadConnection.ConnectedApplication);
                    EnableCADButtons(true);
                    MessageBox.Show($"Successfully connected to {acadConnection.ConnectedApplication}!", 
                                    "Connection Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Failed to connect to AutoCAD or ZWCAD.\n\nPlease ensure one of the following is installed and running:\n• AutoCAD\n• ZWCAD", 
                                    "Connection Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Connection error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DisconnectButton_Click(object sender, RoutedEventArgs e)
        {
            if (acadConnection == null)
            {
                MessageBox.Show("AutoCAD connection is not available.", 
                                "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                acadConnection.Disconnect();
                UpdateConnectionStatus(false, "");
                EnableCADButtons(false);
                MessageBox.Show("Disconnected from AutoCAD/ZWCAD", "Disconnected", 
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Disconnect error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DrawLineButton_Click(object sender, RoutedEventArgs e)
        {
            if (acadConnection == null)
            {
                MessageBox.Show("AutoCAD connection is not available.", 
                                "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // Draw a sample line from (0,0) to (100,100)
                if (acadConnection.DrawLine(0, 0, 100, 100))
                {
                    MessageBox.Show("Line drawn successfully from (0,0) to (100,100)!", 
                                    "Draw Line", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error drawing line: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GetInfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (acadConnection == null)
            {
                MessageBox.Show("AutoCAD connection is not available.", 
                                "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                string info = acadConnection.GetDocumentInfo();
                MessageBox.Show(info, "Document Information", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting document info: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ZoomExtentsButton_Click(object sender, RoutedEventArgs e)
        {
            if (acadConnection == null)
            {
                MessageBox.Show("AutoCAD connection is not available.", 
                                "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                if (acadConnection.ZoomExtents())
                {
                    MessageBox.Show("Zoomed to extents successfully!", "Zoom Extents", 
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error zooming to extents: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SendCommandButton_Click(object sender, RoutedEventArgs e)
        {
            if (acadConnection == null)
            {
                MessageBox.Show("AutoCAD connection is not available.", 
                                "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // Create a simple input dialog
                string command = Microsoft.VisualBasic.Interaction.InputBox(
                    "Enter AutoCAD/ZWCAD command:", "Send Command", "LINE");
                
                if (!string.IsNullOrEmpty(command))
                {
                    if (acadConnection.SendCommand(command))
                    {
                        MessageBox.Show($"Command '{command}' sent successfully!", "Send Command", 
                                        MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error sending command: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateConnectionStatus(bool connected, string applicationName)
        {
            if (connected)
            {
                ConnectionStatusIndicator.Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27AE60")); // Green
                ConnectionStatusText.Text = $"Connected to {applicationName}";
                ConnectionStatusText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27AE60"));
            }
            else
            {
                ConnectionStatusIndicator.Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E74C3C")); // Red
                ConnectionStatusText.Text = "Not Connected";
                ConnectionStatusText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#7F8C8D"));
            }
        }

        private void EnableCADButtons(bool enabled)
        {
            DrawLineButton.IsEnabled = enabled;
            GetInfoButton.IsEnabled = enabled;
            ZoomExtentsButton.IsEnabled = enabled;
            SendCommandButton.IsEnabled = enabled;
            DisconnectButton.IsEnabled = enabled;
            ConnectButton.IsEnabled = !enabled;
        }

        protected override void OnClosed(EventArgs e)
        {
            // Ensure we disconnect when the window is closed
            if (acadConnection != null && acadConnection.IsConnected)
            {
                acadConnection.Disconnect();
            }
            base.OnClosed(e);
        }

        #endregion
    }


}