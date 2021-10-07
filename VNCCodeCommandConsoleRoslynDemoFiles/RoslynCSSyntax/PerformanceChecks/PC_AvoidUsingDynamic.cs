// From Source Code Analysis with Roslyn -

using System;

public class PC_AvoidUsingDynamic
{
    static void Main(string[] args)
    {
        dynamic a = 13;
        dynamic b = 14;
        dynamic c = a + b;
        Console.WriteLine(c);
    }
}
