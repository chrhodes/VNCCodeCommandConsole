// From Source Code Analysis with Roslyn -

public class PC_PreferLiteralsOverEvaluation
{
    public string fun()
    {
        int x = "EDGE".Length;
        string s = "Edge".Substring(1, 4);

        //"234".TryParse();
        return "EDGE".ToLower();

    }

    public string GetRep(string upper)
    {
        return upper.ToLower();
    }
}
