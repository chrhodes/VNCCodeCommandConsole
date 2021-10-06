// From Source Code Analysis with Roslyn -

public class QC_MultipleReturnStatements
{
    int fun(int x)
    {
        x++;
        if (x == 0)
            return x;
        else
            return x + 2;
    }

    double funny3(int x)
    {
        return x / 12;
    }
}