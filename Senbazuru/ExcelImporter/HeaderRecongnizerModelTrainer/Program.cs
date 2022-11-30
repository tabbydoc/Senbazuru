using Senbazuru.HirarchicalExtraction;


namespace HeaderRecongnizerModelTrainer
{
    class Program
    {
        static void Main(string[] args)
        {
            // Model Training
            HirarchicalModel model = new HirarchicalModel();

            model.LoadModel();

            model.HirarchicalModelFileLoading(false);


            //model.Train();

            //model.SaveModel();

            model.Testing();

            model.Evaluation();
        }
    }
}
