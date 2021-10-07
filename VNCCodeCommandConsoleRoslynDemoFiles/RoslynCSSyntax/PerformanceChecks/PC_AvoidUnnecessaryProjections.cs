// From Source Code Analysis with Roslyn -

using System.Collections.Generic;
using System.Linq;

public class PC_AvoidUnnecessaryProjections
{
    void fun(IEnumerable<int> nums)
    {
        var vals = nums.ToList();
        foreach (var v in vals)
        {
        }
    }
}
