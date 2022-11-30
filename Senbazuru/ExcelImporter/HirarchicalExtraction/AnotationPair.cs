using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Senbazuru.HirarchicalExtraction
{
    public class AnotationPair
    {
        public IList<Range> CellList = null;
        public int indexParent = 0;
        public int indexChild = 0;

        // Feature Vector denotes the list of feature values.
        public NodePotentialFeatureVector nodepotentialfeaturevector = null;
        //Feature Vaector of the first hierarhy candidate
        private NodePotentialFeatureVector nodeFeatureAdjacentVector = null;

        public AnotationPair(IList<Range> CellList, int indexParent, int indexChild)
        {
            this.CellList = CellList;
            this.indexParent = indexParent;
            this.indexChild = indexChild;
        }

        public static bool operator ==(AnotationPair pair1, AnotationPair pair2)
        {
            return pair1.indexChild == pair2.indexChild && pair1.indexParent == pair2.indexParent ? true : false;
        }

        public static bool operator !=(AnotationPair pair1, AnotationPair pair2)
        {
            return pair1.indexChild != pair2.indexChild || pair1.indexParent == pair2.indexParent ? true : false;
        }

        //FeaturevectorOfFirstChild is needed for comparison with 
        public bool FeaturevectorOfFirstChildNull()
        {
            if (this.nodeFeatureAdjacentVector is null) return true;
            else return false;
        }
        public NodePotentialFeatureVector featureAdjacentVector
        {
            get
            {
                return this.nodeFeatureAdjacentVector;
            }
            set
            {
                this.nodeFeatureAdjacentVector = value;
            }
        }

    }
}
