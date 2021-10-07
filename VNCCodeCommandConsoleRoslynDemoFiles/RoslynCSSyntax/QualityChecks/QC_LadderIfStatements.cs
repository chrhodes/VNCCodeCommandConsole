// From Source Code Analysis with Roslyn -

public class QC_LadderIfStatements
{
    void IsBad()
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
            f23();
        if (c0() == 24)
            f24();
    }

    void IsGood()
    {
        if (c1() == 13)
            c13();

        if (c2() == 45)
            c45();
    }

    public int c0() { return 0; }
    public int c1() { return 0; }
    public int c2() { return 0; }

    public void c13() { }
    public void c45() { }

    public void f1() { }
    public void f2() { }
    public void f3() { }
    public void f4() { }

    public void f23() { }
    public void f24() { }

}