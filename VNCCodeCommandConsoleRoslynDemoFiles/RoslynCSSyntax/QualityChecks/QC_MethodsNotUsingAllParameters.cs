// From Source Code Analysis with Roslyn -

public class QC_MethodsNotUsingAllParameters
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
    }
}