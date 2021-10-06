// From Source Code Analysis with Roslyn -

public class QC_MagicNumbersUsedInConditions
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
        int[] g = new int[1];

        return g[x];
    }

    bool zun()
    {
        float z = 9;

        if (z > 3.4)
            return false;
        else
            return true;

    }
}