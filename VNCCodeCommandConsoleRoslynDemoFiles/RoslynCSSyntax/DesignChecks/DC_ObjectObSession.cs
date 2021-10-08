// From Source Code Analysis with Roslyn -

using System;
public class DC_ObjectObsession
{
    //Bad: “a” could have been more specifically typed.
    void fun(object a, int x, float d)
    {
    }
    void funny(int x)
    {
    }
    //Bad: This could have used more specific parameter type
    object soFunny(object one)
    {
        return one;
    }
}
