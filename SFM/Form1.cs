using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;


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

        // Pfad zu den Ordnern Fenster1 bis Fenster6
        private string[] folderPaths = new string[6];

        public Form1()
        {
            InitializeComponent();
            CheckIsContent();
            InitializeDashboard();
            // Binden des FormClosing-Ereignisses an die Form1_FormClosing-Methode
            this.FormClosing += Form1_FormClosing;
        }

        private void CheckIsContent()
        {
            // Setze die Pfadvariablen für Fenster1-Fenster6
            string projectDirectory = Directory.GetCurrentDirectory();
            string parentDirectory = Directory.GetParent(projectDirectory).FullName; // Ein Verzeichnis über dem Projektverzeichnis
            string grandParentDirectory = Directory.GetParent(parentDirectory).FullName; // Zwei Verzeichnisse über dem Projektverzeichnis
            string greatGrandParentDirectory = Directory.GetParent(grandParentDirectory).FullName; // Drei Verzeichnisse über dem Projektverzeichnis
            for (int i = 0; i < 6; i++)
            {
                folderPaths[i] = Path.Combine(greatGrandParentDirectory, $"Fenster{i + 1}");
            }

            // Überprüfe, ob mindestens ein Ordner Dateien enthält
            bool anyFolderContainsFiles = false;
            foreach (string folderPath in folderPaths)
            {
                if (Directory.Exists(folderPath) && Directory.GetFiles(folderPath).Length > 0)
                {
                    anyFolderContainsFiles = true;
                    break;
                }
            }

            // Wenn mindestens ein Ordner Dateien enthält, frage den Benutzer, ob das Dashboard leer geöffnet werden soll
            if (anyFolderContainsFiles)
            {
                DialogResult result = MessageBox.Show("Möchten Sie das Dashboard leer öffnen?", "Dashboard öffnen", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    // Lösche den Inhalt der Ordner Fenster1-Fenster6
                    foreach (string folderPath in folderPaths)
                    {
                        DeleteFolderContents(folderPath);
                    }
                }
            }
        }


        // Methode zum Löschen des Inhalts eines Ordners
        private void DeleteFolderContents(string folderPath)
        {
            try
            {
                // Lösche alle Dateien im Ordner
                foreach (string file in Directory.GetFiles(folderPath))
                {
                    File.Delete(file);
                }

                // Lösche alle Unterverzeichnisse und ihre Inhalte
                foreach (string directory in Directory.GetDirectories(folderPath))
                {
                    DeleteFolderContents(directory);
                    Directory.Delete(directory);
                }
            }
            catch (Exception ex)
            {
                // Behandle etwaige Ausnahmen, z.B. Zugriffsverweigerungen
                MessageBox.Show($"Fehler beim Löschen des Ordners {folderPath}: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

                // Kopiere die ausgewählte Datei in den entsprechenden Ordner
                string destinationDirectory = directoryPath;
                string destinationFile = Path.Combine(destinationDirectory, Path.GetFileName(selectedFile));
                File.Copy(selectedFile, destinationFile, true);

                // Erstelle das Fenster mit der ausgewählten Datei
                var windowControl = CreateWindow(destinationFile, x, y, windowWidth, windowHeight);

                // Suche und entferne den "Datei auswählen"-Button, der sich in der Mitte des Fensters befindet
                foreach (Control control in Controls)
                {
                    if (control is System.Windows.Forms.Button && control.Text == "Datei auswählen")
                    {
                        // Überprüfe, ob der Button in der Mitte des aktuellen Fensters liegt
                        if (Math.Abs(control.Location.X + control.Width / 2 - (x + windowWidth / 2)) <= 5 && Math.Abs(control.Location.Y + control.Height / 2 - (y + windowHeight / 2)) <= 5)
                        {
                            Controls.Remove(control);
                            break; // Beende die Schleife, sobald der Button gefunden und entfernt wurde
                        }
                    }
                }

                // Erstelle den Vollbild-Button
                System.Windows.Forms.Button fullscreenButton = new System.Windows.Forms.Button();
                fullscreenButton.Text = "Vollbild"; // Button-Text entsprechend des Index
                fullscreenButton.Size = new Size(windowWidth / 4, 20); // Button-Größe auf ein Viertel der Fensterbreite und Höhe 20 setzen
                System.Drawing.Point buttonLocation = new System.Drawing.Point(x + (windowWidth - fullscreenButton.Width) / 2, y + windowHeight);
                fullscreenButton.Location = buttonLocation;
                fullscreenButton.Click += (sender, e) => ToggleFullscreen(windowControl); // Event-Handler für den Button-Click hinzufügen
                Controls.Add(fullscreenButton); // Button zum Formular hinzufügen
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

                // Überprüfe, ob das Fenster maximiert oder minimiert ist
                if (presentation.Windows[1].WindowState != Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowNormal)
                {
                    presentation.Windows[1].WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowNormal;
                }

                // Ändere die Breite des Fensters
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
                webBrowser.NewWindow += new CancelEventHandler(webBrowser_NewWindow);
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
                webBrowser.NewWindow += new CancelEventHandler(webBrowser_NewWindow);
                Controls.Add(webBrowser);

                windowControl = webBrowser;
            }

            return windowControl;
        }

        private void webBrowser_NewWindow(object sender, CancelEventArgs e)
        {
            e.Cancel = true;
            WebBrowser webBrowser = sender as WebBrowser;
            if (webBrowser != null)
            {
                HtmlElement link = webBrowser.Document.ActiveElement;
                if (link != null && link.TagName == "A")
                {
                    string url = link.GetAttribute("href");
                    webBrowser.Navigate(url);
                }
            }
        }
        private void ToggleFullscreen(Control control)
        {
            if (fullscreenForm == null)
            {
                // Erstelle das Vollbildformular
                fullscreenForm = new Form();
                fullscreenForm.FormBorderStyle = FormBorderStyle.None;
                fullscreenForm.WindowState = FormWindowState.Maximized;

                if (control is WebBrowser)
                {
                    // Erstelle ein Duplikat des WebBrowser-Elements
                    WebBrowser browserDuplicate = new WebBrowser
                    {
                        ScriptErrorsSuppressed = true,
                        Size = new System.Drawing.Size(Screen.PrimaryScreen.WorkingArea.Width, Screen.PrimaryScreen.WorkingArea.Height - 40),
                    };
                    browserDuplicate.Url = ((WebBrowser)control).Url;

                    fullscreenForm.Controls.Add(browserDuplicate);
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

        // Ereignisbehandlungsmethode für das Schließen des Formulars
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Beende alle laufenden Excel-Prozesse
            var excelProcesses = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (var process in excelProcesses)
            {
                process.Kill();
            }

            // Beende alle laufenden PowerPoint-Prozesse
            var powerPointProcesses = System.Diagnostics.Process.GetProcessesByName("POWERPNT");
            foreach (var process in powerPointProcesses)
            {
                process.Kill();
            }
            // Kurze Verzögerung, um sicherzustellen, dass die Prozesse beendet werden
            System.Threading.Thread.Sleep(1000); // Warten Sie 1 Sekunde (1000 Millisekunden)
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

    }
}
