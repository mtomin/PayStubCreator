using Microsoft.Win32;
using System;
using System.Linq;
using System.Windows;
using WinForms = System.Windows.Forms;
using Methods;
using OfficeOpenXml;
using System.IO;

namespace PayStubCreator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        bool errorHappened;
        int numberOfErrors = 0;
        int numberOfSuccessful = 0;
        public MainWindow()
        {
            InitializeComponent();
            
        }

        private void FileAddButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel 97-2003 files(*.xls)|*.xls|Excel 2016 files(*.xlsx)|*.xlsx";

            if (fileDialog.ShowDialog().Value)
            {
                numberOfErrors = 0;
                numberOfSuccessful = 0;
                CreatePayStubFromFile(fileDialog.FileName);
                if (numberOfErrors == 0)
                {
                    string message= "Pay stub successfully created";
                    MessageBox.Show(message, "Operation successful!", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    string logfileName = Path.GetDirectoryName(fileDialog.FileName) + @"\paystub_logfile.log";
                    string message = String.Format("An error was encountered! Please consult the logfile located in {0} for more details", logfileName);
                    MessageBox.Show(message, "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void FolderAddButton_Click(object sender, RoutedEventArgs e)
        {
            string[] acceptableExtensions = { ".xls", ".xlsx" };
            WinForms.FolderBrowserDialog folderDialog = new WinForms.FolderBrowserDialog();

            if (folderDialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                numberOfErrors = 0;
                numberOfSuccessful = 0;
                string[] inputFiles = Directory.GetFiles(folderDialog.SelectedPath, "*.*", SearchOption.TopDirectoryOnly).Where(file => acceptableExtensions.Contains(System.IO.Path.GetExtension(file))).ToArray();

                foreach (string fileName in inputFiles)
                {
                    CreatePayStubFromFile(fileName);
                }
                if (numberOfErrors == 0)
                {
                    string message = String.Format("All {0} pay stubs were successfully created", numberOfSuccessful);
                    MessageBox.Show(message, "Operation successful!", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else if (numberOfSuccessful == 0)
                {
                    string logfileName = folderDialog.SelectedPath + @"\paystub_logfile.log";
                    string message = String.Format("{0} of errors encountered! No pay stubs created. Please consult the logfile located in {0} for more details", logfileName);
                    MessageBox.Show(message, "Operation unsuccessful!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    string logfileName = folderDialog.SelectedPath + @"\paystub_logfile.log";
                    string message = String.Format("{0} pay stubs were successfully generated. {1} errors encountered. Please consult the logfile located in {2} for more details", numberOfSuccessful, numberOfErrors, logfileName);
                    MessageBox.Show(message, "Partial success!", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }
        private void CreatePayStubFromFile(string fileName)
        {
            Company company = new Company();
            Employee employee = new Employee();
            MessageBox.Show(fileName);

            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(fileName)))
            {
                var worksheet = xlPackage.Workbook.Worksheets.First();
                errorHappened = false;
                //delete existing old logfile
                string logfileName = Path.GetDirectoryName(fileName) + @"\paystub_logfile.log";
                if (File.Exists(logfileName))
                    {
                        File.Delete(logfileName);
                    }

                //Create a lambda adapter to pass filename and logfilename to LogError
                //LogErrorAdapter was created because it is not possible to unsubscribe the lambda function
                EventHandler<EventData> LogErrorAdapter = (sender, eventData) => LogError(sender, eventData, fileName, logfileName);
                Readers.ErrorOccured += LogErrorAdapter;
                Readers.ReadCompanyData(company, worksheet);
                
                if (!errorHappened)
                {
                    Readers.GetEmployeeData(employee, worksheet);
                }
                if (!errorHappened)
                {
                    Readers.LoadDBInfo(employee);
                }
                if (!errorHappened)
                {
                    string outputFilePath = System.IO.Path.ChangeExtension(fileName, ".pdf");
                    Writers.ExportToPDF(company, employee, outputFilePath);
                }

                if (!errorHappened)
                {
                    numberOfSuccessful++;
                }
                Readers.ErrorOccured -= LogErrorAdapter;
            }
        }

        private void LogError(object sender, EventData eventData, string fileName, string logfileName)
        {
            eventData.ErrorDescription += string.Format(" {0}. The pay stub for that file was not generated.", fileName);
            using (FileStream fs = File.Open(logfileName, FileMode.Append, FileAccess.Write))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    sw.WriteLine(eventData.ErrorDescription);
                }
            }
            errorHappened = true;
            numberOfErrors++;
        }
    }
}
