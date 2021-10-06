// From Source Code Analysis with Roslyn -

public class QC_MagicNumbersInMathematics
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