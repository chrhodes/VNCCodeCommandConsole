// From Source Code Analysis with Roslyn -

using System;
public class DC_EmptyInterfaces
{
    //        namespace DesignLibrary
    //{
    public interface IGoodInterface
    {
        void funny();
    }

    public interface IBadInterface // Violates rule
    {
    }
    //}
}
