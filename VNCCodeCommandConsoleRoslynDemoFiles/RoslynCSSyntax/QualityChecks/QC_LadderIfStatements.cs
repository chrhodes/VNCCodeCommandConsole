// From Source Code Analysis with Roslyn -

public class QC_LadderIfStatements
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
        if (c0() == 23)
            f22();
        if (c0() == 24)
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

    public int c0() { return 0; }
    public int c1() { return 0; }

    public int c2() { return 0; }

    public void f1() { }
    public void c13() { }
    public void c45() { }
    public void f2() { }
    public void f3() { }
    public void f4() { }
    public void f5() { }
    public void f6() { }
    public void f21() { }
    public void f22() { }
    public void read_that() { }
}