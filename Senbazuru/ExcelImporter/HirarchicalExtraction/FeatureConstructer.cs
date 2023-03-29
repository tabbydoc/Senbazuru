using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;


namespace Senbazuru.HirarchicalExtraction
{
    public class FeatureConstructer
    {
        public List<AnotationPair> anotationPairList;
        public List<AnotationPairEdge> anotationPairEdgeList;

        private Worksheet sheet = null;
        private List<Range> celllist;

        private int RANDOMSAMPLECOUNT = 3;
        private int RANDOMSAMPLEEDGECOUNT = 1;


        /// <summary>
        /// This method is used for construction the feature list for the training data set
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="AttributeRange"></param>
        /// <param name="indexPairList"></param>
        public FeatureConstructer(Worksheet sheet, Range AttributeRange, List<Tuple<int, int>> indexPairList)
        {
            this.AnotationPairConstruction(AttributeRange, sheet, indexPairList);
            this.AnotationPairEdgeConstruction(false);
            this.NodeFeatureVectorConstruction();
            //this.EdgeFeatureVectorConstruction();
            this.sheet = sheet;
        }

        public FeatureConstructer(Worksheet sheet, Range AttributeRange, bool sampleConstruction = false)
        {
            this.AnotationPairConstruction(AttributeRange, sheet, sampleConstruction);
            //Comment for debug heuristic method
            //this.AnotationPairEdgeConstruction(false);
            this.NodeFeatureVectorConstruction();
            //this.EdgeFeatureVectorConstruction();
            this.sheet = sheet;
        }

        private void AnotationPairConstruction(Range AttributeRange, Worksheet sheet, List<Tuple<int, int>> indexPairList)
        {

            this.anotationPairList = new List<AnotationPair>();
            this.celllist = new List<Range>();

            int rowcount = AttributeRange.Rows.Count;
            int colnum = AttributeRange.Column;

            for (int i = 1; i <= rowcount; i++)
            {
                celllist.Add(AttributeRange.Cells[i, colnum]);
            }

            for (int i = 0; i < indexPairList.Count; i++)
            {
                AnotationPair pair = new AnotationPair(celllist, indexPairList[i].Item1, indexPairList[i].Item2);
                anotationPairList.Add(pair);
            }
        }

        private void AnotationPairConstruction(Range AttributeRange, Worksheet sheet, bool SampleConstruction = true)
        {

            this.anotationPairList = new List<AnotationPair>();
            this.celllist = new List<Range>();

            int rowcount = AttributeRange.Rows.Count;
            int colnum = AttributeRange.Column;

            for (int i = 1; i <= rowcount; i++)
            {
                celllist.Add(AttributeRange.Cells[i, colnum]);
            }

            // Exhaust Construction
            if (SampleConstruction)
            {
                for (int i = 0; i < celllist.Count; i++)
                {
                    for (int j = i + 1; j < celllist.Count; j++)
                    {
                        if (i != j)
                        {
                            AnotationPair pair = new AnotationPair(celllist, i, j);
                            anotationPairList.Add(pair);
                        }

                    }
                }
            }
            else
            {
                for (int i = 0; i < celllist.Count; i++)
                {
                    if (celllist.Count - i <= this.RANDOMSAMPLECOUNT) break;
                    for (int j = 0; j < this.RANDOMSAMPLECOUNT; j++)
                    {
                        Random rand = new Random();
                        int index = rand.Next(celllist.Count - i);
                        AnotationPair pair = new AnotationPair(celllist, i, i + index);
                        anotationPairList.Add(pair);
                    }
                }
            }
        }

        private void AnotationPairEdgeConstruction(bool SampleConstruction = true)
        {

            anotationPairEdgeList = new List<AnotationPairEdge>();

            // Exhaust Construction
            if (SampleConstruction)
            {
                for (int i = 0; i < anotationPairList.Count; i++)
                {
                    for (int j = i + 1; j < anotationPairList.Count; j++)
                    {
                        if (i != j)
                        {
                            AnotationPairEdge PairEdge = new AnotationPairEdge(anotationPairList[i], anotationPairList[j]);
                            anotationPairEdgeList.Add(PairEdge);
                        }

                    }
                }
            }
            else
            {
                for (int i = 0; i < anotationPairList.Count; i++)
                {
                    if (anotationPairList.Count - i <= this.RANDOMSAMPLEEDGECOUNT) break;
                    for (int j = 0; j < this.RANDOMSAMPLEEDGECOUNT; j++)
                    {
                        Random rand = new Random();
                        int index = rand.Next(anotationPairList.Count);
                        AnotationPairEdge PairEdge = new AnotationPairEdge(anotationPairList[i], anotationPairList[index]);
                        anotationPairEdgeList.Add(PairEdge);
                    }
                }
            }
        }

        private void NodeFeatureVectorConstruction()
        {
            //Meximum time of operations (debug only)
            TimeSpan findMax(IList<TimeSpan> items) { 
                TimeSpan max = TimeSpan.MinValue;
                foreach (TimeSpan item in items) { 
                    if (item > max)
                        max = item;
                }
                return max;
            }

            ModelFeatures Features = new ModelFeatures();
            DateTime t1;
            DateTime t2;
            List<TimeSpan> t3;
            for (int i = 0; i < anotationPairList.Count; i++)
            {
                //Console.WriteLine(i + "th pair constructed!");
                List<int> featureVector = new List<int>();
                //Console.WriteLine("----");
                t3 = new List<TimeSpan>();
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureAdjacent(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //0
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1= DateTime.Now;
                featureVector.Add(Features.BFeatureBlankCellMiddle(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //1
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureChildindentationGreater(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //2
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureChildindexGreater(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //3
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureChildSizeSmaller(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //4
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureContainColonAndTotal(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //5
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureIndentationLarger(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //6
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureIndentationMiddle(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //7
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureIndentationShorter(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //8
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureParentRoot(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //9
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureStyleAdjacent(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //10
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureBoldDiffer(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //11
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureItalicDiffer(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //12
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureUnderlineDiffer(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //13
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureBackgroundDiffer(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //14
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureParentIsEmptyCell(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //15
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureChildIsEmptyCell(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //16
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureIndentationDifferent(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //17
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureHorisontalAligmentDifferent(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //18
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureVerticalAligmentDifferent(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //19
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                t1 = DateTime.Now;
                featureVector.Add(Features.BFeatureDataTypeDifferent(this.celllist, anotationPairList[i].indexParent, anotationPairList[i].indexChild)); //20
                t2 = DateTime.Now;
                t3.Add(t2 - t1);
                //Output of Debug information
                //Console.WriteLine(String.Format("Max time {0} for feature {1}", findMax(t3), t3.IndexOf(findMax(t3))));

                NodePotentialFeatureVector nodepotentialfeaturevector = new NodePotentialFeatureVector(featureVector);
                anotationPairList[i].nodepotentialfeaturevector = nodepotentialfeaturevector;

                if ((i > 0) && (anotationPairList[i].indexParent == anotationPairList[i - 1].indexParent) &&
                    (anotationPairList[i].nodepotentialfeaturevector.features == anotationPairList[i - 1].nodepotentialfeaturevector.features)) {
                    anotationPairList[i].featureAdjacentVector = anotationPairList[i - 1].featureAdjacentVector;
                }
            }
        }



        private void EdgeFeatureVectorConstruction()
        {
            ModelFeatures Features = new ModelFeatures();

            for (int i = 0; i < anotationPairEdgeList.Count; i++)
            {
                Console.WriteLine(i + "th edge constructed!");
                List<int> featureVector = new List<int>();
                featureVector.Add(Features.EFeatureStylisticAffinity(this.celllist, anotationPairEdgeList[i].pair1.indexParent, anotationPairEdgeList[i].pair1.indexChild, anotationPairEdgeList[i].pair2.indexParent, anotationPairEdgeList[i].pair2.indexChild));
                featureVector.Add(Features.EFeatureMetaDataAffinity(this.celllist, anotationPairEdgeList[i].pair1.indexParent, anotationPairEdgeList[i].pair1.indexChild, anotationPairEdgeList[i].pair2.indexParent, anotationPairEdgeList[i].pair2.indexChild));
                featureVector.Add(Features.EFeatureAdjacentDependency(this.celllist, anotationPairEdgeList[i].pair1.indexParent, anotationPairEdgeList[i].pair1.indexChild, anotationPairEdgeList[i].pair2.indexParent, anotationPairEdgeList[i].pair2.indexChild));
                EdgePotentialFeatureVector edgepotentialfeaturevector = new EdgePotentialFeatureVector(featureVector, this.anotationPairList.IndexOf(anotationPairEdgeList[i].pair1), this.anotationPairList.IndexOf(anotationPairEdgeList[i].pair2));
                anotationPairEdgeList[i].edgepotentialfeaturevector = edgepotentialfeaturevector;
            }
        }
    }
}
