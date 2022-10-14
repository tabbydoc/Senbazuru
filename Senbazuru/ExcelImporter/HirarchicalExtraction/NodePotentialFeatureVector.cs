﻿using System;
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
    }
}
