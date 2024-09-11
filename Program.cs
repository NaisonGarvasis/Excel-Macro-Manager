using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;
using Newtonsoft.Json;

namespace Excel_Macro_Manager
{
    internal class Program
    {
        private static string folderPath = @"D:\RnD\Excel Sample\1";
        private static string folderPathToWriteFiles = @"D:\RnD\Excel Sample\1\";
        
        static void Main(string[] args)
        {
            IList<ExcelConfig> excelConfigList =  InitializeConfig();
            
            foreach (string file in Directory.EnumerateFiles(folderPath, "*.xls*"))
            {
                ProcessFile(file, excelConfigList);
            }
        }


        /// <summary>
        /// Load config file content
        /// </summary>
        /// <returns></returns>
        private static IList<ExcelConfig> InitializeConfig()
        {
            var serializer = new JsonSerializer();
            IList<ExcelConfig> excelConfigList = new List<ExcelConfig>();
            using (var streamReader = new StreamReader("ExcelConfig.json"))
            using (var textReader = new JsonTextReader(streamReader))
            {
                excelConfigList = serializer.Deserialize<List<ExcelConfig>>(textReader);
            }
            return excelConfigList;
        }

        /// <summary>
        /// Process File
        /// </summary>
        /// <param name="filename"></param>
        public static void ProcessFile(string filename, IList<ExcelConfig> excelConfigList)
        {
            Console.WriteLine(" Processing file... " + filename);
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Workbook workbook = excelApp.Workbooks.Open(filename);
            WriteMacroAsTextFile(workbook);
            UpdateMacroCodeModule(workbook,excelConfigList);
            workbook.Save();
            excelApp.Quit();
            Console.WriteLine(" Compted file... " + filename);
        }

        /// <summary>
        /// Update Macro Content
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="excelConfigList"></param>
        static void UpdateMacroCodeModule(Workbook workbook, IList<ExcelConfig> excelConfigList)
        {
            foreach (VBComponent component in workbook.VBProject.VBComponents)
            {
                CodeModule module = component.CodeModule;
                string name = module.Name;
                if (module.Name != "ThisWorkbook" && !module.Name.StartsWith("Sheet"))
                {
                    string lines2 = module.get_Lines(1, module.CountOfLines);
                    string[] lines = module.get_Lines(1, module.CountOfLines).Split(
                            new string[] { "\r\n" },
                            StringSplitOptions.RemoveEmptyEntries);
                    //Console.WriteLine(component.CodeModule.Name);

                    foreach (ExcelConfig config in excelConfigList)
                    {
                        for (int i = 0; i < lines.Length; i++)
                        {
                            if (lines[i].Contains(config.existingPath))
                            {
                                lines[i] = lines[i].Replace(config.existingPath, config.newPath);
                                module.ReplaceLine(i + 1, lines[i]);
                            }
                        }
                    }
                }
            }
        }
        
        static void WriteMacroAsTextFile(Workbook workbook)
        {
            foreach (VBComponent component in workbook.VBProject.VBComponents)
            {
                CodeModule module = component.CodeModule;
                if (module.Name != "ThisWorkbook" && !module.Name.StartsWith("Sheet"))
                {
                    component.Export(folderPathToWriteFiles + component.CodeModule.Name + ".txt");
                    component.Properties.ToString();
                }
            }
        }

      static void RunMacro(Workbook workbook, string macroName)
        {
            try
            {
                workbook.Application.Run(macroName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error running macro '{macroName}': {ex.Message}");
            }
        }
    }
}
