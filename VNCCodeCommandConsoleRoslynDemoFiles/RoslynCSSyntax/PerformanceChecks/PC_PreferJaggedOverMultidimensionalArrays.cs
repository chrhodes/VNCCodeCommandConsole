﻿// From Source Code Analysis with Roslyn -

public class PC_PreferJaggedOverMultidimensionalArrays
{
    int[][] jaggedArray =
    {
            new int[] {1,2,3,4},
            new int[] {5,6,7},
            new int[] {8},
            new int[] {9}
        };

    int[,] multiDimArray =
    {
            {1,2,3,4},
            {5,6,7,0},
            {8,0,0,0},
            {9,0,0,0}
        };
}
