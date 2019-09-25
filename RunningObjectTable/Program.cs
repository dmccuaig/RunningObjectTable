using System;

namespace RunningObjectTable
{
    class Program
    {
        static void Main()
        {
            string state = RotUtil.IsRunning("AutoCAD.Application") ? "is" : "is NOT";
            Console.WriteLine($"AutoCAD {state} running");
            Console.ReadKey();
        }
    }
}
