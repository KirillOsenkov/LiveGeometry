using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DynamicGeometry;

namespace PolylineRouting
{
    class Program
    {
        static void Main(string[] args)
        {
            RoutingAlgorithm algorithm = new RoutingAlgorithmDijkstra();

            string[] input = File.ReadAllLines("obstacle.dat");
            algorithm.ParseInput(input[0], input[1], input[2]);

            var path = algorithm.ShortestRoute();
            var result = path.PolylineLength().Round(2).ToString();

            Console.WriteLine(result);
        }
    }
}
