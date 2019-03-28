using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace CallExcelVbaMacro.Services
{
    public class ExcelService
    {
        public ExcelService()
        {
            Initialize();
        }

        public void RunMacro(string selectedPath, string macroName, string argument1, string argument2,
            string argument3)
        {
            var application = new Application();
            Workbook xlWorkBook = null;
            try
            {
                //selectedPath = "D:\\sources\\SpigaoExcelConnector\\Src\\ConnecteurExcel.xlam";
                xlWorkBook = application.Workbooks.Open(selectedPath);
                xlWorkBook.IsAddin = true;

                //~~> Run the macros by supplying the necessary arguments
                //xlApp.Run("TestMacro", "Hello from C# Client", "Demo to run Excel macros from C#"); //multi args
                //xlApp.Run("TestMacro"); // marche
                //xlApp.Run("Longueil1.xls!TestMacro"); // marche
                //xlApp.Run("Longueil1.xls!Module1.TestMacro"); // marche
                //xlApp.Run("Module1.TestMacro"); // marche
                //xlApp.Run("testcnf");//Works
                //xlApp.Run("TestMsg"); //Works

                //macroArgument1 = "D:\\tests\\Dossiers de tests\\STD\\Longueil1\\Longueil1.xml";
                //macroArgument2 = "True";
                //macroName = "mGetXlsxFromXml";
                application.Run(macroName, argument1, argument2, argument3);
                //xlApp.Run("mGetXlsxFromXml", "D:\\tests\\Dossiers de tests\\STD\\Longueil1\\Longueil1.xml", "True");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw new NullReferenceException(e.Message);
            }
            finally
            {
                xlWorkBook.Close(false);
                application.Quit();

                CleanUp(application, xlWorkBook);
            }
        }

        private void CleanUp(Application xlApp, Workbook xlWorkBook)
        {
            ReleaseObject(xlApp);
            ReleaseObject(xlWorkBook);
        }

        private static void Initialize()
        {
            //var processes = Process.GetProcessesByName("EXCEL");
            //foreach (var process in processes)
            //{
            //    process.Kill();
            //}
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}