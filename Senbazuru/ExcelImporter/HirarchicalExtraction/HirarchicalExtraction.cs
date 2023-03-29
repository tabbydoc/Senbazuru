using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace Senbazuru.HirarchicalExtraction
{
    public class HirarchicalExtraction
    {
        private static String BASE_DIR = Path.GetFullPath("../../../resources/FrameFinder/data/");
        public static String EXCEL_FILES_DIR = Path.Combine(BASE_DIR, "sheets");
        private static String FF_RESULTS_DIR = Path.Combine(BASE_DIR, "results");
        public static string PAIRS_DIR = Path.Combine(BASE_DIR, "pairs");

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
            List<string> parList;  


            var ffData = new Dictionary<string, Range>();

            int tblNum = -1; //-1 value for test only - The one table on the sheet

            AutomaticExtractionModel model = new AutomaticExtractionModel();
            model.LoadModel();
            DirectoryInfo Dir = new DirectoryInfo(EXCEL_FILES_DIR);
            FileInfo[] files = Dir.GetFiles(REGULAREXPRESSION_FILE);
            int cntr = 0;
            foreach (FileInfo file in files)
            {
                //excelReader = new ExcelReaderInterop();
                //excelReader.OpenExcel(file.FullName);
                workbook = excelapp.Workbooks.Open(file.FullName);
                sheet = workbook.Sheets[SHEETNUM];
                Debug.WriteLine(sheet.Name + " processing");
                if (tblNum == -1)
                    ffFile = Path.Combine(FF_RESULTS_DIR, file.Name + "____" + sheet.Name);
                else
                    ffFile = Path.Combine(FF_RESULTS_DIR, file.Name + "____" + sheet.Name + "____" + tblNum);
                ffData = GetFFData(ffFile, sheet);
                if (ffData == null)
                {
                    throw new InvalidDataException("No data about table postiton");
                }

                ModelFeatures modelFeatures = new ModelFeatures();

                constructer = new FeatureConstructer(sheet, ffData["range"], true);
                List<AnotationPair> anotationPairList = constructer.anotationPairList;
                parList = new List<string>();
                if (par["alg"] == "1")
                {
                    //Algorithm 1: Classification
                    int startPoint = 0;
                    for (int i = 0; i < anotationPairList.Count; i++)
                    {
                        if ((startPoint > anotationPairList[i].indexParent) || (startPoint >= anotationPairList[i].indexChild))
                            continue;
                        
                        if ((anotationPairList[i].indexParent == 1) && (anotationPairList[i].indexChild == 2)) {
                           Debug.WriteLine("alarm!");
                        }
                        
                        

                        /*
                        if ((anotationPairList[i].indexParent == 2) && ((anotationPairList[i].indexChild == 25) || (anotationPairList[i].indexChild == 26)))
                            Debug.WriteLine(anotationPairList[i].nodepotentialfeaturevector.FeatureVectorInString());
                        */
                        
                        IList<int> npv = anotationPairList[i].nodepotentialfeaturevector.getFeatures();
                        if ((npv[15] == 1) || (npv[16] == 1)) continue;
                        if (
                            (npv[2] == 1 //Atribute parent's identation greater that Child
                            || npv[4] == 1 //Child’s font size is smaller than parent’s
                            || npv[11] == 1 //One cell has BOLD font and one not
                            || npv[12] == 1 //One cell has ITALIC font and one not
                            || npv[13] == 1 //One cell has UNDERLINE font and one not
                            || npv[14] == 1 //Pair cells have different background
                            || npv[17] == 1 //Pair cells have different identation
                            || npv[18] == 1 //Pair cells have different horisontal aligment
                            || npv[19] == 1 //Pair cells have different vertical aligment
                            || npv[20] == 1 //Pair cells have different datatypes
                            || (npv[5] == 4) //Parent has ":"
                            )
                            //&& ((npv[5] == 0) || (npv[5] == 4))
                            && ((npv[5] != 2))
                            && (npv[15] == 0) //Parent is not empty
                            && (npv[16] == 0) //Child is not empty
                            )
                        // Main features, highlited differences in Parent-Child pair
                        {
                            //if ((npv[0] == 1) ||((npv[0] == 0) && (npv[1] == 1))) //Adjacent cells
                            if (npv[0] == 1) //Adjacent cells
                                {
                                //Potentionally this is a pair
                                anotationPairList[i].nodepotentialfeaturevector.label = true;
                            }
                            else if (
                                    //In our model we do not pay attention on empty cell, just pass them    
                                    //(npv[1] !=2) //There are no middle empty cell
                                    (npv[3] == 1) //Child’s row index is greater than parent’s
                                    //&& (npv[5] != 1) //Has not middle cell containing keywords like "total" or ":" semicolon
                                                     //&& (npv[7] == 0) //There isn't middle cell with indentation between the pair’s
                                    && (
                                        (npv[2] == 1)
                                        || npv[4] == 1 //Child’s font size is smaller than parent’s
                                        || npv[11] == 1 //One cell has BOLD font and one not
                                        || npv[12] == 1 //One cell has ITALIC font and one not
                                        || npv[13] == 1 //One cell has UNDERLINE font and one not
                                        || npv[14] == 1 //Pair cells have different background
                                        || (npv[17] == 1)
                                        || (npv[18] == 1)
                                        || (npv[19] == 1)
                                        || ((npv[20] == 1) || (npv[20] == 2)) //Different or mixed data type of cells
                                        || (npv[5] == 4) //Parent has ":"
                                        )
                                    )
                            {
                                //Compare with adjacent pair
                                for (int j = i - 1; j >= 0; j--)
                                {
                                    //Review pair with this parent only
                                    if (anotationPairList[j].indexParent == anotationPairList[i].indexParent)
                                    {
                                        IList<int> npvj = anotationPairList[j].nodepotentialfeaturevector.getFeatures();
                                       
                                        if (npvj[16] != 1)
                                        {
                                            if (anotationPairList[j].nodepotentialfeaturevector.label == true)
                                            {
                                                if (anotationPairList[i].nodepotentialfeaturevector.similarityOfVectors(npvj) == true)
                                                {
                                                    anotationPairList[i].nodepotentialfeaturevector.label = true;
                                                    break; //The first candidate pair is needed
                                                }
                                            }
                                            else if (anotationPairList[i].nodepotentialfeaturevector.similarityOfVectors(npvj) == true)
                                                break;
                                        }
                                    }
                                    else break;
                                }
                            }



                        }
                        else {
                            //Pair doesn't has any features. If previous pair is labeled start new hierarhy from this level
                            if ((i > 0) && (anotationPairList[i].nodepotentialfeaturevector.equialityOfCellsData()))
                            {
                                startPoint = anotationPairList[i].indexChild;
                            }
                        }
                        //Debug information
                        //if ((anotationPairList[i].indexParent == 10) && (anotationPairList[i].nodepotentialfeaturevector.label))
                        //if ((anotationPairList[i].indexParent == 10) && (anotationPairList[i].indexChild == 11 || anotationPairList[i].indexChild == 14|| anotationPairList[i].indexChild == 17))
                        if (anotationPairList[i].nodepotentialfeaturevector.label)
                        {
                            string s = anotationPairList[i].indexParent + " " + anotationPairList[i].indexChild;
                            parList.Add(s);
                            //if (anotationPairList[i].nodepotentialfeaturevector.label == true) { 
                            //Debug.Write("Parent: " + anotationPairList[i].indexParent + " Child: " + anotationPairList[i].indexChild);
                            //Debug.Write(" - ");
                            //Debug.WriteLine(anotationPairList[i].nodepotentialfeaturevector.FeatureVectorInString() + " " + anotationPairList[i].nodepotentialfeaturevector.label);

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



                if (parList.Count > 0)
                    savePairs2File(parList, sheet.Name + "xls__" + sheet.Name);
                workbook.Close();
                //Save results

                cntr++;
                //if (cntr >3) break;
            }
        }



        static Dictionary<string, Range> GetFFData(string fileName, Worksheet sheet)
        {
            string[] lineFFRes;
            int labelCol = 1;
            string lastPredictLine = "1";
            int startIdx = 0;
            Dictionary<string, Range> result = new Dictionary<string, Range>
            {
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

        static void savePairs2File(List<string> pairs, string fileName, string path="") {
            if (path.Equals(""))
                path = PAIRS_DIR;
            System.IO.File.WriteAllLines(path+"/"+fileName, pairs);



        }


    }
}
