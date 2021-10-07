// From Source Code Analysis with Roslyn -

using System;
public class DC_StaticMembersOnGenericTypes
{
    class A<T>
    {
        public static int fun()
        {
            return 10;
        }
        public int funny<T>()
        {
            return 0;
        }
    }
}
