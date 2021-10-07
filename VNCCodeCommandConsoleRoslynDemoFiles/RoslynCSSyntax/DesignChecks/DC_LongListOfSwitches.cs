// From Source Code Analysis with Roslyn -

using System;
using System.Collections.Generic;
public class DC_LongListOfSwitches
{
    public void fun(int a)
    {
        object nothing = null;
        switch (a)
        {
            case 1: print(1); break;
            case 2: print(5); break;
            case 3: print(4); break;
            case 4: print(2); break;
            case 5: print(8); break;
            case 6: print(7); break;
            case 7: print(7); break;
            default: print(nothing); break;
        }
    }

    public void fun2(int a)
    {
        switch (a + 1)
        {
            case 1: dothat(); break;
            case 2: dothese(); break;
        }
        object nothing = null;
        int g = 0;
        switch (g)
        {
            case 1: print(1); break;
            case 2: print(5); break;
            case 3: print(4); break;
            case 4: print(2); break;
            case 5: print(8); break;
            case 6: print(7); break;
            case 7: print(7); break;
            default: print(nothing); break;
        }
    }

    void print(object o)
    {

    }

    void dothat() { }
    void dothese() { }

}
