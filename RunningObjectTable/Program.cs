using System;

namespace RunningObjectTable
{
    class Program
    {
        static void Main()
        {
            string state = RotUtil.IsRunning("AutoCAD.Application") ? "is" : "is NOT";
            Console.WriteLine($"AutoCAD {state} running\n");


            foreach (var name in RotUtil.GetRunningObjectNames())
            {
                Console.WriteLine(name);
            }

            Console.Write("\nPress a key to continue....");
            Console.ReadKey();
        }
    }
}
