// From Source Code Analysis with Roslyn -

public class QC_MagicNumbersUsedInConditions
{
    int x = 8;

    bool IsValidPW(string password)
    {
        if (password.Length < 5)
            return false;

        if (password.Length < x)
            return false;

        return password.Length >= 7;
    }

    bool IsValidPW2(string password)
    {
        if (5 > password.Length)
            return false;

        return 7 <= password.Length;
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