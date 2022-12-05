using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
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

namespace VerificationNet
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Application.Current.Resources["CurrentFilePath"] = "";

        public class ApplicationContext
        {
            public static string CurrentFilePath { get; set; }
            public static string CountCheckedEmployees { get; set; }
        }

        public MainWindow()
        {
            InitializeComponent();
            OpenFileDialog openedfile = new OpenFileDialog();

            openedfile.Filter = "Excel files | *.xls;*.xlsx;*.xlsm:";
            openedfile.Title = "Usuarios Importados";

            if (openedfile.ShowDialog() == true)
            {
                //Application.Current.Resources["CurrentFilePath"] = openedfile.FileName.ToString();
                ApplicationContext.CurrentFilePath = openedfile.FileName.ToString();
            }
            //Application.Current.Resources["currentfilepath"] = something;
        }

        DataView ImportarTodoArchivoExcel(String ruta)
        {
            //String ruta = "C:\\Users\\te-bautistaa\\Documents\\Importar_Usuarios.xlsx";

            //Crear conexion String
            string Readconexion = string.Format("Provider = Microsoft.Jet.OleDb.4.0; Data Source = {0}; Extended Properties = 'Excel 8.0; HDR=Yes;IMEX=1;MAXSCANROWS=0'", ruta);
            OleDbConnection conector = default(OleDbConnection);
            conector = new OleDbConnection(Readconexion);

            //Abrir Conexion
            conector.Open();

            //Crear consulta
            OleDbCommand consulta = default(OleDbCommand);

            consulta = new OleDbCommand("SELECT * FROM [EMPLEADOS$]", conector);

            //Crear Adaptador
            OleDbDataAdapter adaptador = new OleDbDataAdapter();
            //adaptador.SelectCommand = UpdateConsulta;
            adaptador.SelectCommand = consulta;

            //Crear Dataset
            DataSet ds = new DataSet();

            adaptador.Fill(ds);

            return ds.Tables[0].DefaultView;
        }

        DataView ImportarArchivoExcel(String ruta, string empnumber)
        {
            if (empnumber != "EmployeedataNotMatch")
            {
                //Crear conexion String
                string Readconexion = string.Format("Provider = Microsoft.Jet.OleDb.4.0; Data Source = {0}; Extended Properties = 'Excel 8.0; HDR=Yes;IMEX=1;MAXSCANROWS=0'", ruta);
                OleDbConnection conector = default(OleDbConnection);
                conector = new OleDbConnection(Readconexion);

                conector.Open();
                OleDbCommand ValidateEmployee = new OleDbCommand("SELECT COUNT(*) FROM[EMPLEADOS$] WHERE IdRH = " + empnumber, conector);

                int count = (int)ValidateEmployee.ExecuteScalar();
                //int count = Convert.ToInt32(counter);
                conector.Close();

                if (count > 0)
                {
                    string Writeconexion = string.Format("Provider = Microsoft.Jet.OleDb.4.0; Data Source = {0}; Extended Properties = 'Excel 8.0;'", ruta);
                    OleDbConnection ConectorUpdate = default(OleDbConnection);
                    ConectorUpdate = new OleDbConnection(Writeconexion);

                    ConectorUpdate.Open();

                    OleDbCommand UpdateConsulta = default(OleDbCommand);
                    UpdateConsulta = new OleDbCommand("UPDATE [EMPLEADOS$] SET IsChecked = 'CHECKED' WHERE IdRH =" + empnumber, ConectorUpdate);
                    UpdateConsulta.ExecuteNonQuery();
                    ConectorUpdate.Close();

                    //Abrir Conexion
                    conector.Open();

                    //Crear consulta
                    OleDbCommand consulta = default(OleDbCommand);

                    consulta = new OleDbCommand("SELECT IdRH, Nombre, Departamento, Area, IsChecked FROM [EMPLEADOS$] WHERE IdRH =" + empnumber, conector);

                    OleDbCommand CountCheckedEmployees = new OleDbCommand("SELECT COUNT(*) FROM[EMPLEADOS$] WHERE IsChecked = 'CHECKED'", conector);

                    int countedCheckedEmployees = (int)CountCheckedEmployees.ExecuteScalar();

                    CounterLbl.Content = countedCheckedEmployees.ToString();

                    //Crear Adaptador
                    OleDbDataAdapter adaptador = new OleDbDataAdapter();
                    //adaptador.SelectCommand = UpdateConsulta;
                    adaptador.SelectCommand = consulta;

                    //Crear Dataset
                    DataSet ds = new DataSet();

                    adaptador.Fill(ds);
                    conector.Close();
                    return ds.Tables[0].DefaultView;
                }
                else
                {
                    AlertLabelTxt.Visibility = Visibility.Visible;
                    return null;
                }
            }
            else
            {
                AlertLabelTxt.Visibility = Visibility.Visible;
                return null;
            }

        }
        private string FixEmployeeNumber(string employee)
        {
            int prefix, SumEmployeeCharList;
            List<string> EmployeeCharList = new List<string>();
            string SumValuePrefix;
            string EmployeeNumberwithPrefixRemoved;

            string employeenumberprefix = employee.Substring(0, 3);

            if (employee.Length == 9)
            {
                // for 000000 employees
                SumEmployeeCharList = 0;
                EmployeeNumberwithPrefixRemoved = employee.Remove(0, 3);
                for (int counter = 0; counter < EmployeeNumberwithPrefixRemoved.Length; counter++)
                {
                    EmployeeCharList.Add(EmployeeNumberwithPrefixRemoved[counter].ToString());
                    SumEmployeeCharList = SumEmployeeCharList + int.Parse(EmployeeNumberwithPrefixRemoved[counter].ToString());
                }

                SumValuePrefix = employee.Substring(0, 2);
                if (SumEmployeeCharList == int.Parse(SumValuePrefix))
                {
                    return EmployeeNumberwithPrefixRemoved.TrimStart('0');
                }
                else
                {
                    // for 00000 employees
                    SumEmployeeCharList = 0;
                    EmployeeNumberwithPrefixRemoved = employee.Remove(0, 2);
                    for (int counter = 0; counter < EmployeeNumberwithPrefixRemoved.Length; counter++)
                    {
                        EmployeeCharList.Add(EmployeeNumberwithPrefixRemoved[counter].ToString());
                        SumEmployeeCharList = SumEmployeeCharList + int.Parse(EmployeeNumberwithPrefixRemoved[counter].ToString());
                    }

                    SumValuePrefix = employee.Substring(0, 1);
                    if (SumEmployeeCharList == int.Parse(SumValuePrefix))
                    {
                        return EmployeeNumberwithPrefixRemoved.TrimStart('0');
                    }
                    else
                    {
                        return "EmployeedataNotMatch";
                    }
                }
            }
            if (employee.Length == 5 || employee.Length == 6)
            {
                return employee;
            }
            else
            {
                // for 00000 employees
                SumEmployeeCharList = 0;
                EmployeeNumberwithPrefixRemoved = employee.Remove(0, 3);
                for (int counter = 0; counter < EmployeeNumberwithPrefixRemoved.Length; counter++)
                {
                    EmployeeCharList.Add(EmployeeNumberwithPrefixRemoved[counter].ToString());
                    SumEmployeeCharList = SumEmployeeCharList + int.Parse(EmployeeNumberwithPrefixRemoved[counter].ToString());
                }

                SumValuePrefix = employee.Substring(0, 2);
                if (SumEmployeeCharList == int.Parse(SumValuePrefix))
                {
                    return EmployeeNumberwithPrefixRemoved.TrimStart('0');
                }
                else
                {
                    // for 00000 employees
                    SumEmployeeCharList = 0;
                    EmployeeNumberwithPrefixRemoved = employee.Remove(0, 2);
                    for (int counter = 0; counter < EmployeeNumberwithPrefixRemoved.Length; counter++)
                    {
                        EmployeeCharList.Add(EmployeeNumberwithPrefixRemoved[counter].ToString());
                        SumEmployeeCharList = SumEmployeeCharList + int.Parse(EmployeeNumberwithPrefixRemoved[counter].ToString());
                    }

                    SumValuePrefix = employee.Substring(0, 1);
                    if (SumEmployeeCharList == int.Parse(SumValuePrefix))
                    {
                        return EmployeeNumberwithPrefixRemoved.TrimStart('0');
                    }
                    else
                    {
                        return "EmployeedataNotMatch";
                    }
                }
            }
        }

        private void OnKeyDownHandler(object sender, KeyEventArgs e)      
        {
            if (e.Key == Key.Return)
            {
                AlertLabelTxt.Visibility = Visibility.Hidden;
                String EmployeeNumberWithPrefix = FixEmployeeNumber(txtCredential.Text);

                string XLSPath = ApplicationContext.CurrentFilePath;
                dvgEmployees.ItemsSource = ImportarArchivoExcel(XLSPath, EmployeeNumberWithPrefix);
            }
        }
    }
}
