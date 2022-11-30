using System;
using System.IO;

namespace FrameFinder
{
    class Config
    {
        // this could be modified to absolute path
        public static String BASEDIR = Path.GetFullPath("../../../resources/FrameFinder/");
        //public static String BASEDIR = "E:\\devel\\RSF_2021\\Senbazuru\\Senbazuru\\ExcelImporter\\resources\\FrameFinder\\";
        //"../../../resources/FrameFinder";
        private static String dat = Path.Combine(BASEDIR, "data");
        // directory to store the original spreadsheets
        public static String SHEETDIR = Path.Combine(dat, "sheets");
        // directory to store the output:
        // each spreadsheet labeled with semantic labels for each row
        public static String OUTPUTDIR = Path.Combine(dat, "results");

        // files to store intermediate results
        public static String CRFTEMPDIR = Path.Combine(dat, "tmp");
        public static String CRFTMPFEATURE = Path.Combine(CRFTEMPDIR, "feature.tmp");
        public static String CRFTMPPREDICT = Path.Combine(CRFTEMPDIR, "predict.tmp");

        // template file for CRF++ to parse the provided features
        public static String CRFPPTEMPLATEPATH = Path.Combine(dat, "template");
        // training data
        public static String CRFTRAINDATAPATH = Path.Combine(dat, "saus_train.data");

        /*****************************************
        * please specify the directory of CRF++
        *****************************************/
        // directory of installed CRF++
        public static String CRFPPDIR = Path.Combine(BASEDIR, "CRF++-0.58");
    }
}
