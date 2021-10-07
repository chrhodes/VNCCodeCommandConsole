// From Source Code Analysis with Roslyn -

using System.Collections.Generic;

public class QC_MagicNumbersUsedInIndex
{
    int sad(int b)
    {
        int x = 323;
        Dictionary<int, int> dic = new Dictionary<int, int>();
        int z = dic[x] + x + dic[323] + dic[17];

        return z + b;
    }

    float happy(float c)
    {
        float d = 234;
        Dictionary<float, float> dic = new Dictionary<float, float>();
        float z = dic[d];

        return z;
    }
}