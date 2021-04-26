// From Source Code Analysis with Roslyn -

using System;
using System.Runtime.Serialization;

public class MagicNumbersInMathematics
{
    float fun(int g)
    {
        int size = 10000;
        g += 23456;//bad code. magic number 23456 is used.
        g += size;
        return g / 10;
    }

    decimal updateRate(decimal rate)
    {
        return rate / 0.2345M;
    }

    decimal updateRateM(decimal rateM)
    {
        decimal basis = 0.2345M;
        return rateM / basis;
    }
}

public class MagicNumbersUsedInIndex
{
    int fun(int b)
    {
        int x = 323;
        int z = dic[x] + x + dic[323];
        return z + b;
    }

    float funny(float c)
    {
        int d = 234;
        Dictionary<float, string> dic = getDic();
        float z = dic[d];
        return z;
    }

    Dictionary<float, string> getDic()
    {
        return new Dictionary<float, string>();
    }
}

public class MagicNumbersUsedInConditions
{
    int x = 8;

    bool IsGood(string password)
    {
        if (password.Length < 5)
            return false;
        return password.Length >= 7;
    }

    int fun()
    {
        return g[x];
    }

    bool zun()
    {
        if (z > 3.4)
            return false;
        else
            return true;
    }
}

public class LadderIfStatements
{
    void fun()
    {
        //Call a function only once
        if (c1() == 1)
            f1();
        if (c1() == 2)
            f2();
        if (c1() == 3)
            f3();
        if (c1() == 4)
            f4();
        if (co() == 23)
            f22();
        if (co() == 24)
            f21();
    }

    void funny()
    {
        read_that();
        if (c1() == 3)
            c13();
        if (c2() == 34)
            c45();
    }
}

public class FragmentedConditions
{
    int maybe_do_something(...)
    {
        if (something != -1)
            return 0;
        if (somethingelse != -1)
            return 0;
        if (etc != -1)
            return 0;
        do_something();
    }

    int otherFun()
    {
        if (bailIfIEqualZero == 0)
            return;
        if (string.IsNullOrEmpty(shouldNeverBeEmpty))
            return;
        if (betterNotBeNull == null || betterNotBeNull.RunAwayIfTrue)
            return;
        return 1;
    }
}

public class HungarianNotation
{
    float fIntRate = 4.456;
    float intRate = 4.53;
    long liX = 342;
    bool bCondi = false;
    string name = ""Sam"";
	string strTitle = ""Mr"";
}

public class LotsOfLocalVariablesInMethods
{
    int fun(int x)
    {
        int y = 0;
        x++;
        return x + 1;
    }

    double funny(double x)
    {
        return x / 2.13;
    }
}

public class MethodsNotUsingAllParameters
{
    int fun(int x, int z)
    {
        int y = 0;
        x++;
        return x + 1;
    }

    double funny(double x)
    {
        return x / 2.13;
    }";
}

public class MultipleReturnStatements
{
    int fun(int x)
    {
        x++;
        if (x == 0)
            return x
        else
            return x + 2;
    }

    double funny3(int x)
    {
        return x / 12;
    }
}

public class LongParameterLists
{
    public void f(int a, int b, int c, int d, bool x, bool z, float t)
    {
    }

    public void f3(int a, int b, int c)
    {
    }

    public void f3b(int a, int b, int c, float d, bool x, bool z, float t)
    {
    }

    public void fb(int a, int b, int c)
    {
    }
}

public class GoToLabels
{
    public static void message()
    {
    }

    public void gotoFun()
    {
        // Search:
        for (int i = 0; i < x; i++)
        {
            for (int j = 0; j < y; j++)
            {
                if (array[i, j].Equals(myNumber))
                {
                    goto Found;
                }
            }
        }
    }
    static void Main()
    {
        Console.WriteLine(@""Coffee sizes: 1 = Small 2 = Medium 3 = Large"");
        Console.Write(""Please enter your selection: "");
        string s = Console.ReadLine();
        int n = int.Parse(s);
        int cost = 0;
        switch (n)
        {
            case 1:
                cost += 25;
                break;
            case 2:
                cost += 25;
                goto case 1;
            case 3:
                cost += 50;
                goto case 1;
            default:
                Console.WriteLine(""Invalid selection."");
                break;
        }

        Console.ReadKey();
    }
}

class A
{
    public int g { get; set; }
    public void f1() { }
    public void f2() { }
    public void f3() { }
    public void f4() { }
    public void f5() { }
    public void f6() { }
}

class B
{
    public void f22() { }
    public void f32() { }
}

public class LinesOfCode
{
    void Test()
    {
        for (int i = 1; i < 101; i++)
        {
            if (i % 3 == 0 && i % 5 == 0)
            {
                Console.WriteLine(""FizzBuzz"");
            }
            //Just a comment
            else if (i % 3 == 0)
            {
                Console.WriteLine(""Fizz"");
            }
            else if (i % 5 == 0)
            {
                Console.WriteLine(""Buzz"");
            }
            else
            {
                Console.WriteLine(i);
            }
        }
    }
}

public class ControlFlags
{
    bool cflag = false;

    void update()
    {
        if (!flag)
            if (thatThing())
                flag = true;
    }

    void thatOtherThing()
    {
        bool flag = false;

        if (flag)
        {
            //Do that
        }
        else
        {
            //Do something else
            flag = false;
        }
    }
}

public class CodeToCommentRatio
{
    int fun(int x)
    {
        //update x
        x++;
        return x - 3;
    }

    int add(int x, int y)
    {
    //add these two
    //it might lead to exception
    return x + y;
    }
}


