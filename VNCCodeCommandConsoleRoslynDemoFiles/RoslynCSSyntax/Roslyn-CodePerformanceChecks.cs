// From Source Code Analysis with Roslyn -

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;

public class AvoidBoxing
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

public class AvoidUsingDynamic
{
    static void Main(string[] args)
    {
        dynamic a = 13;
        dynamic b = 14;
        dynamic c = a + b;
        Console.WriteLine(c);
    }
}

public class AvoidExcessiveLocalVariables
{
    void fun()
    {
        int x;
        int y;
        int z;
        int w;
        int ws;
    }
}

public class PreferLiteralsOverEvaluation
{
    public string fun()
    {
        int x = "EDGE".Length;
        string s = "Edge".Substring(1, 4);

        //"234".TryParse();
        return "EDGE".ToLower();

    }

    public string GetRep(string upper)
    {
        return upper.ToLower();
    }
}

public class AvoidVolatileDeclarations
{
    public volatile int i;
    public void Test(int _i)
    {
        i = _i;
    }
}

public class NoObjectArraysInParameters
{
    bool someSearchWithObjParams(params object[] searchTerms)
    {
        return false;
    }

    //bool someSearch(params objectsome[] searchCriteria)
    //{
    //    //Do some search
    //    return true;
    //}

    bool search(int a, int b, params string[] arra)
    {
        return false;
    }

    bool search(string code, int length)
    {
        return true;
    }
}

public class AvoidUnnecessaryProjections
{
    void fun(IEnumerable<int> nums)
    {
        var vals = nums.ToList();
        foreach (var v in vals)
        {
        }
    }
}

public class ValueTypesOverrideEqualsAndGetHashCode
{
    struct Vector : IEquatable<Vector>
    {
        public int X { get; set; }
        public int Y { get; set; }
        public int Z { get; set; }
        public int Magnitude { get; set; }
        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }
            if (obj.GetType() != this.GetType())
            {
                return false;
            }
            return this.Equals((Vector)obj);
        }
        public bool Equals(Vector other)
        {
            return this.X == other.X
            && this.Y == other.Y
            && this.Z == other.Z
            && this.Magnitude == other.Magnitude;
        }
        //Deliberately commented to make this struct a “defaulter”
        //public override int GetHashCode()
        //{
        // return X ^ Y ^ Z ^ Magnitude;
        //}
    }
}

public class AvoidEmptyStringsToFindZeroLength
{
    string s1 = "test";
    public void EqualsTest()
    {
        if (s1 == "")
        {
            Console.WriteLine(@"s1 equals empty
            string.");
        }
    }

    public void LengthTest()
    {
        if (s1 != null && s1.Length == 0)
        {
            Console.WriteLine("s1.Length == 0.");
        }
    }

    public void NullOrEmptyTest()
    {
        if (!String.IsNullOrEmpty(s1))
        {
            Console.WriteLine(@"s1 != null and
            s1.Length != 0.");
        }
    }
}

public class PreferJaggedOverMultidimensionalArrays
{
    int[][] jaggedArray =
    {
            new int[] {1,2,3,4},
            new int[] {5,6,7},
            new int[] {8},
            new int[] {9}
        };

    int[,] multiDimArray =
    {
            {1,2,3,4},
            {5,6,7,0},
            {8,0,0,0},
            {9,0,0,0}
        };
}

public class DoNotReturnArrayFromProperty
{
    string[] nameValues;

    public DoNotReturnArrayFromProperty()
    {
        nameValues = new string[100];
        for (int i = 0; i < 100; i++)
        {
            nameValues[i] = "Sample";
        }
    }

    public string[] Names
    {
        get
        {
            return (string[])nameValues.Clone();
        }
    }

    public static void Main()
    {
        // Using the property in the following manner
        // results in 201 copies of the array.
        // One copy is made each time the loop executes,
        // and one copy is made each time the condition is
        // tested.
        //Test t = new DoNotReturnArrayFromProperty();

        //for (int i = 0; i < t.Names.Length; i++)
        //{
        //    if (t.Names[i] == ("SomeName"))
        //    {
        //        // Perform some operation.
        //    }
        //}
    }

}

public class DoNotUseThreadAbortOrThreadSuspend
{
    public static void Main()
    {
        Thread newThread =
        new Thread(new ThreadStart(TestMethod));
        newThread.Start();
        Thread.Sleep(1000);
        // Abort newThread.
        Console.WriteLine("Main aborting new thread.");
        newThread.Abort("Information from Main.");
        // Wait for the thread to terminate.
        newThread.Join();
        Console.WriteLine(@"New thread terminated – Main exiting.");
    }

    private static void TestMethod()
    {

    }
}