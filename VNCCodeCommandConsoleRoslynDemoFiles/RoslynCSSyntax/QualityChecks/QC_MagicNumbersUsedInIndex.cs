// From Source Code Analysis with Roslyn -

using System.Collections.Generic;
public class QC_MagicNumbersUsedInIndex
{
    int fun(int b)
    {
        int x = 323;
        Dictionary<int, int> dic = new Dictionary<int, int>();
        int z = dic[x] + x + dic[323] + dic[17];
        return z + b;
    }

    float funny(float c)
    {
        float d = 234;
        Dictionary<float, float> dic = getDic();
        float z = dic[d];
        return z;
    }

    Dictionary<float, float> getDic()
    {
        return new Dictionary<float, float>();
    }
}