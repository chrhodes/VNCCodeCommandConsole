// From Source Code Analysis with Roslyn -

using System;
public class DC_TooManyParametersOnGenericTypes
{
    //private static T FromString<T>(string s) where T : struct
    //{
    //    if (typeof(T).Equals(typeof(decimal)))
    //    {
    //        var x = (decimal)System.Convert.ToInt32(s) / 100;
    //        return (T)Convert.ChangeType(x, typeof(T));
    //    }

    //    if (typeof(T).Equals(typeof(int)))
    //    {
    //        var x = System.Convert.ToInt32(s);
    //        return (T)Convert.ChangeType(x, typeof(T));
    //    }

    //    if (typeof(T).Equals(typeof(DateTime)))
    //    {
    //        var x = System.Convert.ToInt32(s);
    //        return (T)Convert.ChangeType(x, typeof(T));
    //    }

    //    return typeof(T);
    //}

    //public string name()
    //{
    //    return string.Empty;
    //}

    ////really a bad idea. Trust me!
    //private static T FromString2<T, T, T>(string s) where T : struct
    //{
    //    if (typeof(T).Equals(typeof(decimal)))
    //    {
    //        var x = (decimal)System.Convert.ToInt32(s) / 100;
    //        return (T)Convert.ChangeType(x, typeof(T));
    //    }
    //    if (typeof(T).Equals(typeof(int)))
    //    {
    //        var x = System.Convert.ToInt32(s);
    //        return (T)Convert.ChangeType(x, typeof(T));
    //    }
    //    if (typeof(T).Equals(typeof(DateTime)))
    //    {
    //    }
    //}
}
