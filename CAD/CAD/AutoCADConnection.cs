using System;
using System.Runtime.InteropServices;
using System.Windows;

namespace CAD
{
    /// <summary>
    /// Provides connection and basic operations with AutoCAD/ZWCAD through COM automation
    /// </summary>
    public class AutoCADConnection
    {
        private dynamic acadApp = null;
        private dynamic acadDoc = null;
        private bool isConnected = false;
        private string connectedApplication = "";

        /// <summary>
        /// Gets whether the connection to AutoCAD/ZWCAD is active
        /// </summary>
        public bool IsConnected => isConnected;

        /// <summary>
        /// Gets the name of the connected application (AutoCAD or ZWCAD)
        /// </summary>
        public string ConnectedApplication => connectedApplication;

        /// <summary>
        /// Attempts to connect to AutoCAD or ZWCAD
        /// </summary>
        /// <returns>True if connection successful, false otherwise</returns>
        public bool Connect()
        {
            try
            {
                // First try to connect to existing AutoCAD instance
                if (TryConnectToAutoCAD())
                {
                    return true;
                }

                // Then try to connect to existing ZWCAD instance
                if (TryConnectToZWCAD())
                {
                    return true;
                }

                // If no existing instance, try to start AutoCAD
                if (TryStartAutoCAD())
                {
                    return true;
                }

                // Finally try to start ZWCAD
                if (TryStartZWCAD())
                {
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Connection error: {ex.Message}", "AutoCAD/ZWCAD Connection", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        private bool TryConnectToAutoCAD()
        {
            // Try version-specific ProgIDs that are actually registered on this system
            string[] progIds = {
                "AutoCAD.Application.25",   // AutoCAD 2025 (found in registry)
                "AutoCAD.Application.24.1", // AutoCAD 2024.1 (found in registry)
                "AutoCAD.Application.24",   // AutoCAD 2024 (found in registry)
                "AutoCAD.Application"       // Generic fallback (found in registry)
            };

            foreach (string progId in progIds)
            {
                try
                {
                    acadApp = Marshal.GetActiveObject(progId);
                    if (acadApp != null)
                    {
                        // Wait a bit for AutoCAD to be fully loaded
                        System.Threading.Thread.Sleep(1000);
                        
                        acadDoc = acadApp.ActiveDocument;
                        isConnected = true;
                        connectedApplication = $"AutoCAD ({progId})";
                        return true;
                    }
                }
                catch (COMException ex)
                {
                    // Check if it's just unavailable (still loading) vs not found
                    const uint MK_E_UNAVAILABLE = 0x800401e3;
                    if ((uint)ex.ErrorCode == MK_E_UNAVAILABLE)
                    {
                        // AutoCAD is starting up, wait and retry
                        System.Threading.Thread.Sleep(2000);
                        try
                        {
                            acadApp = Marshal.GetActiveObject(progId);
                            if (acadApp != null)
                            {
                                acadDoc = acadApp.ActiveDocument;
                                isConnected = true;
                                connectedApplication = $"AutoCAD ({progId})";
                                return true;
                            }
                        }
                        catch
                        {
                            // Continue to next ProgID
                        }
                    }
                }
                catch
                {
                    // Continue to next ProgID
                }
            }
            return false;
        }

        private bool TryConnectToZWCAD()
        {
            try
            {
                acadApp = Marshal.GetActiveObject("ZwCAD.Application");
                if (acadApp != null)
                {
                    acadDoc = acadApp.ActiveDocument;
                    isConnected = true;
                    connectedApplication = "ZWCAD";
                    return true;
                }
            }
            catch
            {
                // ZWCAD not running or not available
            }
            return false;
        }

        private bool TryStartAutoCAD()
        {
            // Try version-specific ProgIDs that are actually registered on this system
            string[] progIds = {
                "AutoCAD.Application.25",   // AutoCAD 2025 (found in registry)
                "AutoCAD.Application.24.1", // AutoCAD 2024.1 (found in registry)
                "AutoCAD.Application.24",   // AutoCAD 2024 (found in registry)
                "AutoCAD.Application"       // Generic fallback (found in registry)
            };

            foreach (string progId in progIds)
            {
                try
                {
                    Type acadType = Type.GetTypeFromProgID(progId);
                    if (acadType != null)
                    {
                        acadApp = Activator.CreateInstance(acadType);
                        acadApp.Visible = true;
                        
                        // Wait for AutoCAD to fully start
                        System.Threading.Thread.Sleep(3000);
                        
                        acadDoc = acadApp.ActiveDocument;
                        isConnected = true;
                        connectedApplication = $"AutoCAD ({progId})";
                        return true;
                    }
                }
                catch
                {
                    // Continue to next ProgID
                }
            }
            return false;
        }

        private bool TryStartZWCAD()
        {
            try
            {
                Type zwcadType = Type.GetTypeFromProgID("ZwCAD.Application");
                if (zwcadType != null)
                {
                    acadApp = Activator.CreateInstance(zwcadType);
                    acadApp.Visible = true;
                    acadDoc = acadApp.ActiveDocument;
                    isConnected = true;
                    connectedApplication = "ZWCAD";
                    return true;
                }
            }
            catch
            {
                // ZWCAD not installed or cannot start
            }
            return false;
        }

        /// <summary>
        /// Disconnects from AutoCAD/ZWCAD
        /// </summary>
        public void Disconnect()
        {
            try
            {
                if (acadDoc != null)
                {
                    Marshal.ReleaseComObject(acadDoc);
                    acadDoc = null;
                }

                if (acadApp != null)
                {
                    Marshal.ReleaseComObject(acadApp);
                    acadApp = null;
                }

                isConnected = false;
                connectedApplication = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Disconnect error: {ex.Message}", "AutoCAD/ZWCAD Connection", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// Draws a line in AutoCAD/ZWCAD
        /// </summary>
        /// <param name="startX">Start X coordinate</param>
        /// <param name="startY">Start Y coordinate</param>
        /// <param name="endX">End X coordinate</param>
        /// <param name="endY">End Y coordinate</param>
        /// <returns>True if successful</returns>
        public bool DrawLine(double startX, double startY, double endX, double endY)
        {
            if (!isConnected || acadDoc == null)
            {
                MessageBox.Show("Not connected to AutoCAD/ZWCAD", "Connection Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            try
            {
                // Create line coordinates array
                double[] startPoint = { startX, startY, 0.0 };
                double[] endPoint = { endX, endY, 0.0 };

                // Get model space
                dynamic modelSpace = acadDoc.ModelSpace;

                // Add line to model space
                dynamic lineObj = modelSpace.AddLine(startPoint, endPoint);

                // Update the display
                acadApp.ZoomExtents();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error drawing line: {ex.Message}", "Drawing Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        /// <summary>
        /// Sends a command to AutoCAD/ZWCAD
        /// </summary>
        /// <param name="command">Command to send</param>
        /// <returns>True if successful</returns>
        public bool SendCommand(string command)
        {
            if (!isConnected || acadDoc == null)
            {
                MessageBox.Show("Not connected to AutoCAD/ZWCAD", "Connection Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            try
            {
                acadDoc.SendCommand(command + "\n");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error sending command: {ex.Message}", "Command Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        /// <summary>
        /// Gets information about the current document
        /// </summary>
        /// <returns>Document information string</returns>
        public string GetDocumentInfo()
        {
            if (!isConnected || acadDoc == null)
            {
                return "Not connected to AutoCAD/ZWCAD";
            }

            try
            {
                string docName = acadDoc.Name;
                string appName = acadApp.Name;
                string version = acadApp.Version;

                return $"Application: {appName}\nVersion: {version}\nDocument: {docName}";
            }
            catch (Exception ex)
            {
                return $"Error getting document info: {ex.Message}";
            }
        }

        /// <summary>
        /// Zooms to extents in AutoCAD/ZWCAD
        /// </summary>
        /// <returns>True if successful</returns>
        public bool ZoomExtents()
        {
            if (!isConnected || acadApp == null)
            {
                MessageBox.Show("Not connected to AutoCAD/ZWCAD", "Connection Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            try
            {
                acadApp.ZoomExtents();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error zooming to extents: {ex.Message}", "Zoom Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        /// <summary>
        /// Destructor to ensure COM objects are released
        /// </summary>
        ~AutoCADConnection()
        {
            Disconnect();
        }
    }
}