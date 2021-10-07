// From Source Code Analysis with Roslyn -

using System;
using System.Drawing;

public class PC_AvoidBoxing
{
    public void fun()
    {
        int x = 32;
        Point p = new Point(10, 10);
        object box = p;
        p.X = 20;
        Console.Write(((Point)box).X);
        object o = x;
    }
}
