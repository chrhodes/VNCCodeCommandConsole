// From Source Code Analysis with Roslyn -

using System;
using System.Runtime.Serialization;

public class DeeplyNestedLoops
{
    void fun2(int x)
    {
        for (int i = 0; i < 10; i++)
        {
            for (int j = 0; j < 10; j++)
                list.Add(i + j);
        }
    }

    void fun(int x)
    {
        for (int i = 0; i < 10; i++)
        {
            for (int j = 0; j < 10; j++)
            {
                for (int k = 2; k < 20; k++)
                    list.Add(i + j + k);
            }
        }
    }

    void straightLoop()
    {
        for (int j = 0; j < 10; j++)
            doThat(j);
    }

    void loopingTheLoopWhile()
    {
        while (true)
            for (int x = 0; x < 10; x++)
                foreach (var z in z[x])
                    doSome(z);
    }

    void loopingTheLoop()
    {
        foreach (var m in newItems)
            foreach (var z in oldItems)
                for (int i = 0; i < z.Items.Count; i++)
                    doThat(i, z, m);
    }

    void fun4(int x)
    {
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

public class OutParameters
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

public class DeeplyNestedIfBlocks
{
    void check(int x)
    {
        if (x < 10)
            doSomeThing();
    }

    void fun2(int x, int y)
    {
        if (x < y)
            if (x + y < 20)
                doThat();
    }

    void fun(int x)
    {
        //really stupid example
        //but you shall get the point
        int x = 20;
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

public class LongListOfSwitches
{
    public void fun(int a)
    {
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
}

public class DataClasses
{
    class A
    {
        void fun() { }
        void fun1(int a) { }
        void fun2(int a, int a) { }
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

public class LocalClasses
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

public class RefusedBequest
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

public class LotsOfMethodOverloads
{
    class DocumentHome
    {
        public Document createDocument(String name)
        {
            // just calls another method with default value
            // of its parameter
            return createDocument(name, -1);
        }

        public Document createDocument(String name, int minPagesCount)
        {
            // just calls another method with default value of its parameter
            return createDocument(name, minPagesCount, false);
        }

        public Document createDocument(String name, int minPagesCount, boolean firstPageBlank)
        {
            // just calls another method with default value of its parameter
            return createDocument(name, minPagesCount, false, "");
        }

        public Document createDocument(String name, int minPagesCount, boolean firstPageBlank, String title)
        {
            // here the real work gets done
        }
    }

    public class EmptyInterfaces
    {
        namespace DesignLibrary
{
    public interface IGoodInterface
    {
        void funny();
    }
    public interface IBadInterface // Violates rule
    {
    }
}
    }

    public class TooManyParametersOnGenericTypes
{
    private static T FromString<T>(string s) where T : struct
    {
        if (typeof(T).Equals(typeof(decimal)))
        {
            var x = (decimal)System.Convert.ToInt32(s) / 100;
            return (T)Convert.ChangeType(x, typeof(T));
        }
        if (typeof(T).Equals(typeof(int)))
        {
            var x = System.Convert.ToInt32(s);
            return (T)Convert.ChangeType(x, typeof(T));
        }
        if (typeof(T).Equals(typeof(DateTime)))
        {
        }
        public string name()
        {
            return string.Empty;
        }

        //really a bad idea. Trust me!
        private static T FromString2<T, T, T>(string s) where T : struct
        {
            if (typeof(T).Equals(typeof(decimal)))
            {
                var x = (decimal)System.Convert.ToInt32(s) / 100;
                return (T)Convert.ChangeType(x, typeof(T));
            }
            if (typeof(T).Equals(typeof(int)))
            {
                var x = System.Convert.ToInt32(s);
                return (T)Convert.ChangeType(x, typeof(T));
            }
            if (typeof(T).Equals(typeof(DateTime)))
            {
            }
        }
    }

    public class StaticMembersOnGenericTypes
    {
        class A<T>
        {
            public static int fun()
            {
                return 10;
            }
            public int funny<T>()
            {
                return 0;
            }
        }
    }

    public abstract class BadAbstractClassWithConstructor
    {
        // Violates rule: AbstractTypesShouldNotHaveConstructors.
        public BadAbstractClassWithConstructor()
        {
            // Add constructor logic here.
        }
        public abstract void fun() { }
    }

    public abstract class GoodAbstractClassWithConstructor
    {
        protected GoodAbstractClassWithConstructor()
        {
            // Add constructor logic here.
        }
    }


        public sealed class SealedClassAndProtectedMembers
        {
            protected void ProtectedMethod() { }
        }


    public class ObjectObSession
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