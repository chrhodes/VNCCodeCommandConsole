// From Source Code Analysis with Roslyn -

using System;
using System.Collections.Generic;
public class DC_DeeplyNestedLoops
{
    void fun2(int x)
    {
        List<int> list = new List<int>();

        for (int i = 0; i < 10; i++)
        {
            for (int j = 0; j < 10; j++)
                list.Add(i + j);
        }
    }

    void fun(int x)
    {
        List<int> list = new List<int>();

        for (int i = 0; i < 10; i++)
        {
            for (int j = 0; j < 10; j++)
            {
                for (int k = 2; k < 20; k++)
                    list.Add(i + j + k);
            }
        }
    }

    private void doThat(int j)
    {

    }

    void straightLoop()
    {
        for (int j = 0; j < 10; j++)
            doThat(j);
    }

    private void doSome(int z)
    {

    }
    void loopingTheLoopWhile()
    {
        List<int> Zzz = new List<int>();

        while (true)
            for (int x = 0; x < 10; x++)
                foreach (var z in Zzz)
                    doSome(z);
    }

    void loopingTheLoop()
    {
        IEnumerable<object> newItems = null;
        IEnumerable<object> oldItems = null;

        foreach (var m in newItems)
            foreach (var z in oldItems)
                for (int i = 0; i < 7; i++)
                    doThat(i, z, m);
    }

    private void doThat(int i, object z, object m)
    {

    }

    void fun4(int x)
    {
        List<int> list = new List<int>();

        for (int m = 0; m < 10; m += 2)
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    for (int k = 2; k < 20; k++)
                        list.Add(i + j + k);
                }
            }
    }
}
