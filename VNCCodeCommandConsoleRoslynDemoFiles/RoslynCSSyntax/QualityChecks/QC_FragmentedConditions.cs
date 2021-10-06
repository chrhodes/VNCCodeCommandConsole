// From Source Code Analysis with Roslyn -

public class QC_FragmentedConditions
{
    int maybe_do_something(int something, int somethingelse, int etc)
    {
        if (something != -1)
            return 0;
        if (somethingelse != -1)
            return 0;
        if (etc != -1)
            return 0;
        do_something();
        return 1;
    }

    public void do_something() { }

    int otherFun()
    {
        int bailIfIEqualZero = 1;
        string shouldNeverBeEmpty = "";

        if (bailIfIEqualZero == 0)
            return 0;
        if (string.IsNullOrEmpty(shouldNeverBeEmpty))
            return 0;

        object betterNotBeNull = null;
        object RunAwayIfTrue = null;

        if (betterNotBeNull == null || betterNotBeNull == RunAwayIfTrue)
            return 0;

        return 1;
    }
}