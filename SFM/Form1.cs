using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace SFM
{
    public partial class Form1 : Form
    {

        private Form fullscreenForm;
        // Erstelle ein Dictionary zur Verfolgung der Zuordnung zwischen Fenstern und Buttons
        Dictionary<Control, System.Windows.Forms.Button> windowButtonMap = new Dictionary<Control, System.Windows.Forms.Button>();


        [DllImport("user32.dll", SetLastError = true)] 
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        public Form1()
        {
            InitializeComponent();
            InitializeDashboard();
        }

        private void InitializeDashboard()
        {
            // Setze die FormWindowState auf maximiert, um sicherzustellen, dass das Formular die Größe des Bildschirms einnimmt
            this.WindowState = FormWindowState.Maximized;

            // Bestimme die Anzahl der Spalten und Zeilen
            int numCols = 3;
            int numRows = 2;

            // Berechne die Größe der Fenster basierend auf der Anzahl der Spalten und Zeilen
            int windowHeight = (Screen.PrimaryScreen.WorkingArea.Height - 60) / numRows; // Höhe jedes Fensters ist (Bildschirmhöhe - 40) / Anzahl der Zeilen, um Platz für die Buttons zu lassen
            int windowWidth = (Screen.PrimaryScreen.WorkingArea.Width - (numCols - 1) * 20) / numCols; // Breite jedes Fensters ist (Bildschirmbreite - (Anzahl der Lücken)) / Anzahl der Spalten

            string projectDirectory = Directory.GetCurrentDirectory();
            string parentDirectory = Directory.GetParent(projectDirectory).FullName; // Ein Verzeichnis über dem Projektverzeichnis
            string grandParentDirectory = Directory.GetParent(parentDirectory).FullName; // Zwei Verzeichnisse über dem Projektverzeichnis
            string greatGrandParentDirectory = Directory.GetParent(grandParentDirectory).FullName; // Drei Verzeichnisse über dem Projektverzeichnis


            // Durchsuche sechs Ordner und erstelle entsprechende Fenster
            for (int i = 1; i <= 6; i++)
            {
                // Berechne die Zeile und Spalte des aktuellen Fensters
                int row = (i - 1) / numCols;
                int col = (i - 1) % numCols;

                // Berechne die Position des Fensters
                int x = col * (windowWidth + 20); // X-Position des Fensters mit 20 Pixel Abstand
                int y = row * (windowHeight + 20); // Y-Position des Fensters mit 20 Pixel Abstand
                string directoryPath = Path.Combine(greatGrandParentDirectory, $"Fenster{i}");

                if (Directory.Exists(directoryPath))
                {
                    string[] files = Directory.GetFiles(directoryPath);
                    if (files.Length > 0)
                    {
                        // Dateien gefunden, erstelle das Fenster
                        var windowControl = CreateWindow(files[0], x, y, windowWidth, windowHeight);

                        // Erstelle den Vollbild-Button
                        System.Windows.Forms.Button button = new System.Windows.Forms.Button();
                        button.Text = "Vollbild"; // Button-Text entsprechend des Index
                        button.Size = new Size(windowWidth / 4, 20); // Button-Größe auf ein Viertel der Fensterbreite und Höhe 20 setzen
                        System.Drawing.Point buttonLocation = new System.Drawing.Point(x + (windowWidth - button.Width) / 2, y + windowHeight);
                        button.Location = buttonLocation;
                        button.Click += (sender, e) => ToggleFullscreen(windowControl); // Event-Handler für den Button-Click hinzufügen
                        Controls.Add(button); // Button zum Formular hinzufügen

                        // Füge die Zuordnung zwischen Fenster und Button zum Dictionary hinzu
                        windowButtonMap.Add(windowControl, button);
                    }
                    else
                    {
                        // Keine Dateien gefunden, erstelle einen Button in der Mitte des Fensters
                        System.Windows.Forms.Button noFileButton = new System.Windows.Forms.Button();
                        noFileButton.Text = "Datei auswählen";
                        noFileButton.Size = new Size(150, 30);
                        System.Drawing.Point buttonLocation = new System.Drawing.Point(x + (windowWidth - noFileButton.Width) / 2, y + (windowHeight - noFileButton.Height) / 2);
                        noFileButton.Location = buttonLocation;
                        noFileButton.Click += (sender, e) => SelectFileAndDisplayWindow(directoryPath, x, y, windowWidth, windowHeight); // Übergebe die erforderlichen Argumente
                        Controls.Add(noFileButton); // Button zum Formular hinzufügen
                    }
                }
            }
        }

        private void SelectFileAndDisplayWindow(string directoryPath, int x, int y, int windowWidth, int windowHeight)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Alle Dateien|*.*"; // Du kannst den Dateifilter anpassen, um bestimmte Dateitypen zu ermöglichen

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Eine Datei wurde ausgewählt
                string selectedFile = openFileDialog.FileName;

                // Erstelle das Fenster mit der ausgewählten Datei
                var windowControl = CreateWindow(selectedFile, x, y, windowWidth, windowHeight);

                // Erstelle den Vollbild-Button
                System.Windows.Forms.Button button = new System.Windows.Forms.Button();
                button.Text = "Vollbild"; // Button-Text entsprechend des Index
                button.Size = new Size(windowWidth / 4, 20); // Button-Größe auf ein Viertel der Fensterbreite und Höhe 20 setzen
                System.Drawing.Point buttonLocation = new System.Drawing.Point(x + (windowWidth - button.Width) / 2, y + windowHeight);
                button.Location = buttonLocation;
                button.Click += (sender, e) => ToggleFullscreen(windowControl); // Event-Handler für den Button-Click hinzufügen
                Controls.Add(button); // Button zum Formular hinzufügen

                // Suche und entferne den "Datei auswählen"-Button
                foreach (Control control in Controls)
                {
                    if (control is System.Windows.Forms.Button && control.Text == "Datei auswählen")
                    {
                        Controls.Remove(control);
                        break; // Beende die Schleife, sobald der Button gefunden und entfernt wurde
                    }
                }

                // Füge die Zuordnung zwischen Fenster und Button zum Dictionary hinzu
                windowButtonMap.Add(windowControl, button);
            }
        }
        private Control CreateWindow(string url, int x, int y, int width, int height)
        {
            Control windowControl = null;

            if (url.EndsWith(".xlsx"))
            {
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                var workbook = excelApp.Workbooks.Open(url);
                excelApp.Visible = true;

                workbook.Windows[1].Width = width;
                workbook.Windows[1].Height = height;

                IntPtr hwnd = new IntPtr(excelApp.Hwnd);

                Panel panel = new Panel
                {
                    Size = new System.Drawing.Size(width, height),
                    Location = new System.Drawing.Point(x, y)
                };

                Controls.Add(panel);

                MoveWindow(hwnd, 0, 0, width, height, true);

                SetParent(hwnd, panel.Handle);

                windowControl = panel;
            }
            else if (url.EndsWith(".pptx") || url.EndsWith(".ppt"))
            {
                Microsoft.Office.Interop.PowerPoint.Application powerPointApp = new Microsoft.Office.Interop.PowerPoint.Application();
                var presentation = powerPointApp.Presentations.Open(url);
                presentation.Windows[1].Width = width;
                presentation.Windows[1].Height = height;

                IntPtr powerPointAppHandle = new IntPtr(powerPointApp.HWND);

                Panel panel = new Panel
                {
                    Size = new System.Drawing.Size(width, height),
                    Location = new System.Drawing.Point(x, y)
                };

                Controls.Add(panel);

                MoveWindow(powerPointAppHandle, 0, 0, width, height, false);

                SetParent(powerPointAppHandle, panel.Handle);

                windowControl = panel;
            }
            else if (url.EndsWith(".pdf") || url.EndsWith(".html"))
            {
                WebBrowser webBrowser = new WebBrowser
                {
                    ScriptErrorsSuppressed = true,
                    Size = new System.Drawing.Size(width, height),
                    Location = new System.Drawing.Point(x, y)
                };

                webBrowser.Navigate(url);
                Controls.Add(webBrowser);

                windowControl = webBrowser;
            }
            else
            {
                WebBrowser webBrowser = new WebBrowser
                {
                    ScriptErrorsSuppressed = true,
                    Size = new System.Drawing.Size(width, height),
                    Location = new System.Drawing.Point(x, y)
                };

                webBrowser.Navigate(url);
                Controls.Add(webBrowser);

                windowControl = webBrowser;
            }

            return windowControl;
        }

        private void ToggleFullscreen(Control control)
        {
            if (fullscreenForm == null)
            {
                // Erstelle das Vollbildformular
                fullscreenForm = new Form();
                fullscreenForm.FormBorderStyle = FormBorderStyle.None;
                fullscreenForm.WindowState = FormWindowState.Maximized;

                // Erstelle ein Duplikat des Steuerelements entsprechend seines Typs
                if (control is WebBrowser)
                {
                    WebBrowser browserDuplicate = new WebBrowser
                    {
                        ScriptErrorsSuppressed = true,
                        Size = new System.Drawing.Size(Screen.PrimaryScreen.WorkingArea.Width, Screen.PrimaryScreen.WorkingArea.Height - 40),
                    };
                    //browserDuplicate.Height = Screen.PrimaryScreen.WorkingArea.Height - 40;
                    //browserDuplicate.Width = Screen.PrimaryScreen.WorkingArea.Width;
                    browserDuplicate.Url = ((WebBrowser)control).Url;

                    fullscreenForm.Controls.Add(browserDuplicate);
                }
                else if (control is Panel)
                {
                    Panel panelDuplicate = new Panel();
                    panelDuplicate.Size = ((Panel)control).Size;
                    panelDuplicate.Location = System.Drawing.Point.Empty;

                    foreach (Control ctrl in ((Panel)control).Controls)
                    {
                        Control ctrlDuplicate = null;
                        if (ctrl is WebBrowser)
                        {
                            WebBrowser webBrowser = new WebBrowser();
                            webBrowser.Size = ctrl.Size;
                            webBrowser.Url = ((WebBrowser)ctrl).Url;
                            ctrlDuplicate = webBrowser;
                        }
                        else
                        {
                            Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)ctrl;
                            Microsoft.Office.Interop.Excel.Application excelApp = workbook.Application;
                            excelApp.Visible = true;
                            workbook.Windows[1].WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;

                            // Füge die Excel-Anwendung selbst zum Panel hinzu
                            IntPtr hwnd = new IntPtr(excelApp.Hwnd);
                            SetParent(hwnd, panelDuplicate.Handle);
                            continue;
                        }
 
                    }

                    fullscreenForm.Controls.Add(panelDuplicate);
                }

                // Zeige das Vollbildformular an
                fullscreenForm.Show();

                // Erstelle den Button zum Beenden des Vollbildmodus
                System.Windows.Forms.Button exitFullscreenButton = new System.Windows.Forms.Button();
                exitFullscreenButton.Text = "Exit Fullscreen";
                exitFullscreenButton.Size = new Size(100, 30);
                System.Drawing.Point buttonLocation = new System.Drawing.Point(20, fullscreenForm.Height - 50);
                exitFullscreenButton.Location = buttonLocation;
                exitFullscreenButton.Click += (sender, e) => ExitFullscreen();
                fullscreenForm.Controls.Add(exitFullscreenButton);
            }
            else
            {
                ExitFullscreen();

                // Füge das ursprüngliche Steuerelement wieder zum ursprünglichen Dashboard hinzu
                control.Parent = this;
                CenterControlInParent(control);
            }
        }

        private void ExitFullscreen()
        {
            if (fullscreenForm != null)
            {
                fullscreenForm.Close();
                fullscreenForm = null;
            }
        }
        private void CenterControlInParent(Control control)
        {
            // Bestimme die Größe des übergeordneten Panels (des kleinen Fensterbereichs)
            int parentWidth = this.Width;
            int parentHeight = this.Height;

            // Verwende die Größe des übergeordneten Panels, um das Steuerelement in seiner Mitte zu positionieren
            int x = (parentWidth - control.Width) / 2;
            int y = (parentHeight - control.Height) / 2;

            // Setze die Position des Steuerelements
            System.Drawing.Point controlLocation = new System.Drawing.Point(x, y);
            control.Location = controlLocation;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Code, der beim Laden des Formulars ausgeführt werden soll
        }
    }
}
