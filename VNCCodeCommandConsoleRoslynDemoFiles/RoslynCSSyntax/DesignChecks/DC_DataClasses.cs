// From Source Code Analysis with Roslyn -

using System;
using System.Collections.Generic;
public class DC_DataClasses
{
    class A
    {
        void fun() { }
        void fun1(int a) { }
        void fun2(int a, int b) { }
        public int Age { get; set; }
        public string Name { get; set; }
    }

    //B is a data class smell as it has only public data
    class B
    {
        public int Age { get; set; }
        public string Name { get; set; }
    }

    //C is also a data class smell as it has only public data
    class C
    {
        public double RateOfInterest { get; set; }
    }
}
