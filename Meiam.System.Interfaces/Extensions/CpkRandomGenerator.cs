using System;
using System.Collections.Generic;
using System.Linq;
using MathNet.Numerics.Distributions;

namespace Meiam.System.Interfaces.Extensions
{
    internal class CpkRandomGenerator
    {
        //private const int SampleSize = 32;
        //private const double TargetCpk = 1.33;

        public static List<double> GenerateCpkSamples(double target, double tolerance,int SampleSize, double TargetCpk)
        {
            var random = new Random();
            var normal = new Normal(target, CalculateSigma(tolerance, TargetCpk));

            List<double> results = new List<double>(SampleSize);
            for (int i = 0; i < SampleSize; i++)
            {
                results.Add(normal.Sample());
            }
            return results;
        }

        private static double CalculateSigma(double tolerance, double cpk)
        {
            return tolerance / (3 * cpk);
        }

        public static double[] GenerateValues(double target, double upperTol, double lowerTol, int count, double minCpk)
        {
            Random rand = new Random();
            double usl = target + upperTol;
            double lsl = target - lowerTol;
            double allowableSigma = (usl - lsl) / (6 * minCpk);

            // 预计算控制参数
            double mean = target;
            double sigma = allowableSigma * 0.9; // 预留10%余量

            // 单次生成合格数据
            while (true)
            {
                var values = Enumerable.Range(0, count)
                    .Select(_ => BoxMullerTransform(rand, mean, sigma))
                    .ToArray();

                if (CalculateCpk(values, usl, lsl) >= minCpk)
                    return values;
            }
        }

        static double BoxMullerTransform(Random rand, double mean, double sigma)
        {
            double u1 = 1.0 - rand.NextDouble();
            double u2 = 1.0 - rand.NextDouble();
            double randStdNormal = Math.Sqrt(-2.0 * Math.Log(u1)) *
                                  Math.Sin(2.0 * Math.PI * u2);
            return mean + sigma * randStdNormal;
        }

        static double CalculateCpk(double[] values, double usl, double lsl)
        {
            double mean = values.Average();
            double sigma = Math.Sqrt(values.Sum(v => Math.Pow(v - mean, 2)) / values.Length);
            double cpu = (usl - mean) / (3 * sigma);
            double cpl = (mean - lsl) / (3 * sigma);
            return Math.Min(cpu, cpl);
        }
    }
}
