﻿using System.Collections.Generic;

namespace Senbazuru.HirarchicalExtraction
{
    public class NodePotentialFeatureVector
    {
        public IList<int> features = new List<int>();
        public bool label = false;

        public NodePotentialFeatureVector(IList<int> features, bool label = false)
        {
            this.features = features;
            this.label = label;
        }

        public IList<int> getFeatures()
        {
            return features;
        }

        //Add for Debug: Feature list in string
        public string FeatureVectorInString()
        {
            return string.Join("; ", features);
        }
        //Compare feature vectors
        public static bool operator ==(NodePotentialFeatureVector features1, NodePotentialFeatureVector features2)
        {
            //Ignore adjacent feature
            if ((features1 == null) && (features2 == null))
                return true;
            if (((features1 == null) && (features2 != null)) || ((features1 != null) && (features2 == null)))
                return false;

            for (int i = 1; i < features1.getFeatures().Count - 1; i++)
            {
                if (features1.getFeatures()[i] != features2.getFeatures()[i])
                    return false;
            }
            return true;
        }

        public static bool operator !=(NodePotentialFeatureVector features1, NodePotentialFeatureVector features2)
        {
            //Ignore adjacent feature
            if ((features1 == null) && (features2 == null))
                return false;
            if (((features1 == null) && (features2 != null)) || ((features1 != null) && (features2 == null)))
                return false;

            for (int i = 1; i < features1.getFeatures().Count - 1; i++)
            {
                if (features1.getFeatures()[i] == features2.getFeatures()[i])
                    return false;
            }
            return true;
        }

        //2 pairs similarity
        public bool similarityOfVectors(IList<int> npv) {
            if ((npv != null) && (features.Count == npv.Count)
                && (features[2] == npv[2]) //Similar identation type
                && (features[3] == npv[3]) //ChildindexGreater
                && (features[4] == npv[4]) //Similar child font size smaler
                && (features[8] == npv[8]) //Similar Identation shoter
                && (features[11] == npv[11]) //Similar font bold
                && (features[12] == npv[12]) //Similar italic
                && (features[13] == npv[13]) //Similar underline
                && (features[14] == npv[14]) //Similar backgorund
                && (features[17] == npv[17]) //Similar identation
                && (features[18] == npv[18]) //Similar horisontal aligment
                )
            {

                return true;
            }


            return false;
        }
        public bool equialityOfCellsData() {
            int cnt = features.Count;
            for (int i = 1; i < cnt; i++) { 
             if (((i != 3) && (features[i] != 0)) || ((i == 3) && (features[i] != 1 ))) return false;
            }

            return true;
        }


  
    }
}
