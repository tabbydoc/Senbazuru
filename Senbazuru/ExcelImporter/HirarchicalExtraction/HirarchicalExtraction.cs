using System;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;

namespace Senbazuru.HirarchicalExtraction
{
    public class HirarchicalExtraction
    {
        private static String BASE_DIR = Path.GetFullPath("../../../resources/FrameFinder/data/");
        public static String EXCEL_FILES_DIR = Path.Combine(BASE_DIR, "sheets");
        private static String FF_RESULTS_DIR = Path.Combine(BASE_DIR, "results");

        private static String REGULAREXPRESSION_FILE = "*.xls*";
        private string startCell;

        static void Main(string[] args)
        {
            int SHEETNUM = 1; //SHEETNUM temporary for testing - only 1-st sheet
            
            //Application APP = new Application();
            //ExcelReaderInterop excelReader;
            Application excelapp = new Application();
            Workbook workbook;
            Worksheet sheet;
            Range startCell;
            Range endCell;
            Range range;
            String ffFile;
            FeatureConstructer constructer;
            Dictionary<string, string> par = Utils.NamedParams(args);
            


            var ffData = new Dictionary<string, Range>();
            
            int tblNum = 0; //0 value for test only - The one table on the sheet

            AutomaticExtractionModel model = new AutomaticExtractionModel();
            model.LoadModel();
            DirectoryInfo Dir = new DirectoryInfo(EXCEL_FILES_DIR);
            FileInfo[] files = Dir.GetFiles(REGULAREXPRESSION_FILE);

            foreach (FileInfo file in files)
            {
                //excelReader = new ExcelReaderInterop();
                //excelReader.OpenExcel(file.FullName);
                workbook = excelapp.Workbooks.Open(file.FullName);
                sheet = workbook.Sheets[SHEETNUM];
                ffFile = Path.Combine(FF_RESULTS_DIR, file.Name + "____" + sheet.Name + "____" + tblNum);
                ffData = GetFFData(ffFile, sheet);
                ModelFeatures modelFeatures = new ModelFeatures();  
                
                
                constructer = new FeatureConstructer(sheet, ffData["range"], true);
                List<AnotationPair> anotationPairList = constructer.anotationPairList;
                if (par["alg"] == "1") {
                    //Algorithm 1: Classification
                    for (int i = 0; i < anotationPairList.Count; i++)
                    {

                        IList<int> npv = anotationPairList[i].nodepotentialfeaturevector.getFeatures();
                        if ((npv[2] == 1 || npv[11] == 1 || npv[12] == 1 || npv[13] == 1 || npv[14] == 1 ) && (npv[3] == 1))
                            anotationPairList[i].nodepotentialfeaturevector.label = true;
                        if (anotationPairList[i].nodepotentialfeaturevector.label == true) { 
                            Debug.Write("Parent: " + anotationPairList[i].indexParent + " Child: " + anotationPairList[i].indexChild);
                            Debug.Write(" - ");
                            Debug.WriteLine(anotationPairList[i].nodepotentialfeaturevector.FeatureVectorInString());
                        }
                    }
                    
                    
                }
                /*
                List<AnotationPair> anotationPairList = constructer.anotationPairList;
                List<AnotationPairEdge> anotationPairEdges = constructer.anotationPairEdgeList;
                IList<NodePotentialFeatureVector> nodepotentialfeaturevector = new List<NodePotentialFeatureVector>();



                for (int i = 0; i < anotationPairList.Count; i++)
                {
                    anotationPairList[i].nodepotentialfeaturevector.label = true;

                }


                model.Testing(anotationPairList, anotationPairEdges);
                AttributeTree attributeTree = model.GetTree(anotationPairList);
                attributeTree.AttributeRelationDictionary(anotationPairList);
                
                foreach (int key in attributeTree.AttributeChildtoParent.Keys) {
                    int child=0;
                    attributeTree.AttributeChildtoParent.TryGetValue(key, out child);
                    Debug.WriteLine(key + ", " + child);
                }
                */



                workbook.Close();
            }
        }

        static Dictionary<string, Range> GetFFData(string fileName, Worksheet sheet)
        {
            string[] lineFFRes;
            int labelCol = 1;
            string lastPredictLine="1";
            int startIdx = 0;
            Dictionary<string, Range> result = new Dictionary<string, Range> {
                ["startCell"] = null,
                ["endCell"] = null,
                ["range"] = null
            };

            using (StreamReader sr = File.OpenText(fileName))
            {
                string s = String.Empty;
                while ((s = sr.ReadLine()) != null)
                {
                    lineFFRes = s.Split('\t');
                    if ((lineFFRes[1].Trim()).Equals("Blank"))
                            continue;
                    if ((result["startCell"] == null) && ((lineFFRes[1].Trim()).Equals("Data"))) 
                    {
                        //Write info about initial data cell
                        result["startCell"] = sheet.Cells[lineFFRes[0], labelCol];
                     
                    }
                    if (result["startCell"] != null) 
                        {
                        if ((lineFFRes[1].Trim()).Equals("Data"))
                            result["endCell"] = sheet.Cells[lineFFRes[0], labelCol];
                        else
                            break;

                    }

                }
                result["range"] = sheet.get_Range(result["startCell"].get_Address() + ":" + result["endCell"].get_Address());

            }
            return result;

        }


    }
}
