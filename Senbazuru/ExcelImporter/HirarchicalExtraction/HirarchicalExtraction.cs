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
                NodePotentialFeatureVector npfv = null;
                
                constructer = new FeatureConstructer(sheet, ffData["range"], true);
                List<AnotationPair> anotationPairList = constructer.anotationPairList;
                if (par["alg"] == "1") {
                    //Algorithm 1: Classification
                    for (int i = 0; i < anotationPairList.Count; i++)
                    {

                        IList<int> npv = anotationPairList[i].nodepotentialfeaturevector.getFeatures();
                        if (anotationPairList[i].FeaturevectorOfFirstChild != null)
                            npfv = anotationPairList[i].nodepotentialfeaturevector;

                        if (
                            (npv[2] == 1 //Atribute parent's identation greater that Child
                            || npv[4]== 1 //Child’s font size is smaller than parent’s
                            || npv[11] == 1 //One cell has BOLD font and one not
                            || npv[12] == 1 //One cell has ITALIC font and one not
                            || npv[13] == 1 //One cell has UNDERLINE font and one not
                            || npv[14] == 1 //Pair cells have different background
                            
                            ) 
                                && (npv[1] == 0) //There are no middle empty cell
                                && (npv[3] == 1) //Child’s row index is greater than parent’s
                                && npv[5] == 0 //Has not middle cell containing keywords like "total" or ":" semicolon
                                && npv[6] == 0 //Has not middle cell with indentation larger the pair’s
                                && npv[7] == 0
                                && npv[8] == 0 //Has not middle cell with indentation between the pair’s
                                && (npv[15] == 0) //Parent is not empty
                                && (npv[16] == 0) //Child is not empty

                            )
                            anotationPairList[i].nodepotentialfeaturevector.label = true;
                        //Debug information
                        if ((anotationPairList[i].indexParent == 10) && (anotationPairList[i].indexChild ==11 || anotationPairList[i].indexChild == 14)) { 
                        //if (anotationPairList[i].nodepotentialfeaturevector.label == true) { 
                            Debug.Write("Parent: " + anotationPairList[i].indexParent + " Child: " + anotationPairList[i].indexChild);
                            Debug.Write(" - ");
                            Debug.WriteLine(anotationPairList[i].nodepotentialfeaturevector.FeatureVectorInString() + " " + anotationPairList[i].nodepotentialfeaturevector.label);
                            
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
