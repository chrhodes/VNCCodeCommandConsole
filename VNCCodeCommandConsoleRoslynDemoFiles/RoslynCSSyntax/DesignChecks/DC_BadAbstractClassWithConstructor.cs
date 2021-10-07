// From Source Code Analysis with Roslyn -

using System;
public abstract class DC_BadAbstractClassWithConstructor
{
    // Violates rule: AbstractTypesShouldNotHaveConstructors.
    public DC_BadAbstractClassWithConstructor()
    {
        // Add constructor logic here.
    }
    //public abstract void fun() { }
}
