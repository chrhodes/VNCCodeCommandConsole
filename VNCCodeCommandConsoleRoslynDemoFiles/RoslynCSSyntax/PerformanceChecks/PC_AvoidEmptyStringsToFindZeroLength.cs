// From Source Code Analysis with Roslyn -

using System;

public class PC_AvoidEmptyStringsToFindZeroLength
{
    string s1 = "test";
    public void EqualsTest()
    {
        if (s1 == "")
        {
            Console.WriteLine(@"s1 equals empty
            string.");
        }
    }

    public void LengthTest()
    {
        if (s1 != null && s1.Length == 0)
        {
            Console.WriteLine("s1.Length == 0.");
        }
    }

    public void NullOrEmptyTest()
    {
        if (!String.IsNullOrEmpty(s1))
        {
            Console.WriteLine(@"s1 != null and
            s1.Length != 0.");
        }
    }
}
