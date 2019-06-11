//AssociationRule.cs

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace APRIORI_SIEUTHI
{
    public class AssociationRule
    {
        #region Properties

        public Itemset X { get; set; }
        public Itemset Y { get; set; }
        public double Support { get; set; }
        public double Confidence { get; set; }

        #endregion

        #region Constructors

        public AssociationRule()
        {
            X = new Itemset();
            Y = new Itemset();
            Support = 0.0;
            Confidence = 0.0;
        }

        #endregion

        #region Methods

        public override string ToString()
        {
            return (X + "-" + Y + ":" + Math.Round(Support, 2) + ":" + Math.Round(Confidence, 2));
        }

        //public override string ToString()
        //{
        //    return (X + " => " + Y + " (support: " + Math.Round(Support, 2) + "%, confidence: " + Math.Round(Confidence, 2) + "%)");
        //}

        //public string ToLuat()
        //{
        //    return (X + " => " + Y + ":" + Math.Round(Support, 2) + ":" + Math.Round(Confidence, 2));
        //}
        #endregion
    }
}
