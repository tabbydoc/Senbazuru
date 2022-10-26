using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Senbazuru.HirarchicalExtraction
{
    public class NodePotentialFeatureVector
    {
        public IList<int> features = new List<int>();
        public bool label = false;

        public NodePotentialFeatureVector(IList<int> features, bool label = false)
        {
            this.features = features ;
            this.label = label ;
        }

        public IList<int> getFeatures() { 
            return features;
        }

        //Add for Debug: Feature list in string
        public string FeatureVectorInString()
        {
            return string.Join("; ", features);
        }
        //Compare feature vectors
        public static bool operator == (NodePotentialFeatureVector fv1, NodePotentialFeatureVector fv2) {
            //Ignore adjacent feature
            for (int i = 1; i < fv1.getFeatures().Count-1; i++)
            {
                if (fv1.getFeatures()[i] != fv2.getFeatures()[i])
                    return false;
            }
            return true;
        }

        public static bool operator !=(NodePotentialFeatureVector features1, NodePotentialFeatureVector features2)
        {
            //Ignore adjacent feature
            for (int i = 1; i < features1.getFeatures().Count - 1; i++)
            {
                if (features1.getFeatures()[i] == features2.getFeatures()[i])
                    return false;
            }
            return true;
        }
    }
}
