// From Source Code Analysis with Roslyn -

public class PC_AvoidVolatileDeclarations
{
    public volatile int i;
    public void Test(int _i)
    {
        i = _i;
    }
}
