// From Source Code Analysis with Roslyn -

using System;
using System.Collections.Generic;
public class DC_RefusedBequest
{
    interface ISomething
    {
        public void f1();
        public void f2();
    }

    interface ISomeOtherThing
    {
        public void other1();
        public void other2();
    }

    class A : ISomething, ISomeOtherThing
    {
        public void other1()
        {
        }
        public void other2()
        {
        }
        public void f1()
        {
            throw new NotImplementedException();
        }
        public void f2()
        {
            throw new NotImplementedException();
        }
    }

    class B : ISomething
    {
        int f_1 = 0;
        int f_2 = 1;
        public void f1()
        {
            Console.WriteLine(f_1);
        }
        public void f2()
        {
            Console.WriteLine(f_2);
        }
    }
}
