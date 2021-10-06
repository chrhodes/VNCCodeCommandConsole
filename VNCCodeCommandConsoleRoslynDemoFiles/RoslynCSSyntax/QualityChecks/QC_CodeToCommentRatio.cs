// From Source Code Analysis with Roslyn -

public class QC_CodeToCommentRatio
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