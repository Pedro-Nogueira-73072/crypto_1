using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Python.Runtime;
using static System.Net.Mime.MediaTypeNames;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace cripto_2
{
    class Program
    {
        static void Encrypt(string scriptName)
        {
            Runtime.PythonDLL = "C:\\Users\\Pedro\\AppData\\Local\\Programs\\Python\\Python312\\python312.dll";
            PythonEngine.Initialize();
            using (Py.GIL())
            {
                dynamic sys = Py.Import("sys");
                sys.path.append("C:\\Users\\Pedro\\Documents\\Visual Studio 2022\\Projects\\Cripto\\cripto_1");

                var pythonScript = Py.Import(scriptName);
                var input_file_path = new PyString("C:\\Users\\Pedro\\Documents\\Visual Studio 2022\\Projects\\Cripto\\cripto_1\\test.xls");
                pythonScript.InvokeMethod("encrypt_excel_file", new PyObject[] { input_file_path });
            }
        }

        static void Decrypt(string scriptName)
        {
            Runtime.PythonDLL = "C:\\Users\\Pedro\\AppData\\Local\\Programs\\Python\\Python312\\python312.dll";
            PythonEngine.Initialize();
            using (Py.GIL())
            {
                dynamic sys = Py.Import("sys");
                sys.path.append("C:\\Users\\Pedro\\Documents\\Visual Studio 2022\\Projects\\Cripto\\cripto_1");

                var pythonScript = Py.Import(scriptName);
                var input_file_path = new PyString("C:\\Users\\Pedro\\Documents\\Visual Studio 2022\\Projects\\Cripto\\cripto_1\\test.xls");
                var cipher_file_path = new PyString("C:\\Users\\Pedro\\Documents\\Visual Studio 2022\\Projects\\Cripto\\cripto_1\\cipher_file");
                pythonScript.InvokeMethod("decrypt_excel_file", new PyObject[] { input_file_path, cipher_file_path });
            }
        }

        static void Main(string[] args)
        {
            //Encrypt("functions");
            Decrypt("functions");
        }
    }
}
