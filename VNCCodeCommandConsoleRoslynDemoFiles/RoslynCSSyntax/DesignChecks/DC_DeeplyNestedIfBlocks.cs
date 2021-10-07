// From Source Code Analysis with Roslyn -

using System;
using System.Collections.Generic;
public class DC_DeeplyNestedIfBlocks
{
    private void doSomeThing()
    {

    }

    void check(int x)
    {
        if (x < 10)
            doSomeThing();
    }

    private void doThat()
    {

    }

    void fun2(int x, int y)
    {
        if (x < y)
            if (x + y < 20)
                doThat();
    }

    private void doOther()
    {

    }

    void fun(int x)
    {
        //really stupid example
        //but you shall get the point

        //Nesting Level 1
        if (x < 10)
        {
            //Nesting Level 2
            if (x - 1 < 10)
            {
                if (x - 2 < 10)//Nesting Level 3
                {
                    if (x - 4 < 10)//Nesting Level 4
                        doThat();
                    else
                        doOther();
                }
            }
        }
    }
}
