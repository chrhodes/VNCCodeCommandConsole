// From Source Code Analysis with Roslyn -

using System;
using System.Collections.Generic;
public class DC_LocalClasses
{
    class A
    {
        //Bad to have local classes.
        class localA
        {
        }
    }
    class B
    {
        void fun()
        {
        }
    }
    class C
    {
        void funny()
        {
        }
    }
}
