using System;
using System.Collections.Generic;

namespace FrameFinder
{
    public class FrameFinder
    {
        /*
         * This method shows how to use the frame finder.
         */
        public static void ProcessEachTable()
        {
            CRFRunner crfRunner = new CRFRunner();
            crfRunner.Train();
            PredictSheetRow predict = new PredictSheetRow();
            foreach (Tuple<string, string, MSheet> eachTable in predict.ScanEachExcel())
            {
                string workbookName = eachTable.Item1;
                string sheetName = eachTable.Item2;
                MSheet mSheet = eachTable.Item3;
                crfRunner.Predict(workbookName, sheetName);
                TransformOutput.Run(workbookName, sheetName, mSheet);
                //PrintLabel(mSheet.Labels);
                HorizontalHierarchyExtractor hhe = new HorizontalHierarchyExtractor();
                hhe.ExtractFromMSheet(mSheet);


            }
        }

        public static void PrintLabel(Dictionary<int, RowLabel> labels)
        {
            foreach (int row in labels.Keys)
            {
                Console.WriteLine("{0} {1}", row, labels[row].ToString());
            }
        }
    }
}
