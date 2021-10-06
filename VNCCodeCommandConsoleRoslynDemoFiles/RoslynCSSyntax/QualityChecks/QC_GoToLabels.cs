// From Source Code Analysis with Roslyn -

using System;

public class QC_GoToLabels
{
    public static void message()
    {
    }

    public void gotoFun()
    {
        // Search:
        int x = 5;
        int y = 5;
        int myNumber = 0;

        int[,] array = new int[2, 3];

        for (int i = 0; i < x; i++)
        {
            for (int j = 0; j < y; j++)
            {
                if (array[i, j].Equals(myNumber))
                {
                    goto Found;
                }
            }
        }

    Found:
        ;

    }

    static void Main()
    {
        Console.WriteLine(@"Coffee sizes: 1 = Small 2 = Medium 3 = Large");
        Console.Write("Please enter your selection: ");
        string s = Console.ReadLine();
        int n = int.Parse(s);
        int cost = 0;
        switch (n)
        {
            case 1:
                cost += 25;
                break;
            case 2:
                cost += 25;
                goto case 1;
            case 3:
                cost += 50;
                goto case 1;
            default:
                Console.WriteLine("Invalid selection.");
                break;
        }

        Console.ReadKey();
    }
}