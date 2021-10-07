// From Source Code Analysis with Roslyn -

using System;
using System.Collections.Generic;
public class DC_OutParameters
{
    static void Method(out int i)
    {
        i = 44;
    }
    static void Main()
    {
        int value;
        Method(out value);
        // value is now 44
    }
}
