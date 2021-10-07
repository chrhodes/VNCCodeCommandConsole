// From Source Code Analysis with Roslyn -

public class PC_NoObjectArraysInParameters
{
    bool someSearchWithObjParams(params object[] searchTerms)
    {
        return false;
    }

    //bool someSearch(params objectsome[] searchCriteria)
    //{
    //    //Do some search
    //    return true;
    //}

    bool search(int a, int b, params string[] arra)
    {
        return false;
    }

    bool search(string code, int length)
    {
        return true;
    }
}
