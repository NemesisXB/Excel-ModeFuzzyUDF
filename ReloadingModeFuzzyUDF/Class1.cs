using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ReloadingModeFuzzyUDF
{
    public class Spread
    {
        public double min { get; set; }
        public double max { get; set; }
        public int count { get; set; }
    }

    public static class ReloadingFunctions
    {
        [ExcelFunction(Name = "MODE.FUZZY", Description = "Calculates the mode of a dataset with a specified \"spread\"", Category = "ReloadingCalculator")]
        public static double[] ModeFuzzy(double[] range, double spreadValue, double spreadIncrement)
        {
            double[] output = new double[3];
            var foos = new List<double>(range);
            foos.RemoveAll(i => i == 0);
            range = foos.ToArray();
            for (int i = 0; i < range.Count(); i++)
            {
                range[i] = Math.Round(range[i], 2);
            }
            double smallest = Math.Round(range.Min(), 2);
            double largest = Math.Round(range.Max(), 2);
            List<Spread> spreadCountList = new List<Spread>();
            Spread spreadSingle = new Spread();
            spreadSingle.min = Math.Round(smallest, 2);
            spreadSingle.max = Math.Round((Math.Min((smallest + spreadValue), largest)) - spreadIncrement, 2);
            spreadSingle.count = 0;
            
            while (spreadSingle.max <= largest)
            {
                for (int i = 0; i < range.Count(); i++)
                {
                    if ((range[i] >= spreadSingle.min) && (range[i] <= spreadSingle.max))
                    {
                        spreadSingle.count++;
                    }
                }
                spreadCountList.Add(spreadSingle);
                spreadSingle = new Spread();
                spreadSingle.min = Math.Round(spreadCountList[spreadCountList.Count - 1].min + spreadIncrement, 2);
                spreadSingle.max = Math.Round(spreadSingle.min + spreadValue - spreadIncrement, 2);
                spreadSingle.count = 0;
            }

            int indexMax
            = !spreadCountList.Any() ? -1 :
            spreadCountList
            .Select((value, index) => new { Value = value, Index = index })
            .Aggregate((a, b) => (a.Value.count > b.Value.count) ? a : b)
            .Index;

            if (indexMax >= 0)
            {
                output[0] = spreadCountList[indexMax].count;
                output[1] = spreadCountList[indexMax].min;
                output[2] = spreadCountList[indexMax].max;
            }
            else
            {
                output[0] = 0;
                output[1] = 0;
                output[2] = 0;
            }
            

            return output;
        }
    }
}
