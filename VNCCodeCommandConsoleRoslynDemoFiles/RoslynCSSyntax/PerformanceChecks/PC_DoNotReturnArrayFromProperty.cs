// From Source Code Analysis with Roslyn -

public class PC_DoNotReturnArrayFromProperty
{
    string[] nameValues;

    public PC_DoNotReturnArrayFromProperty()
    {
        nameValues = new string[100];
        for (int i = 0; i < 100; i++)
        {
            nameValues[i] = "Sample";
        }
    }

    public string[] Names
    {
        get
        {
            return (string[])nameValues.Clone();
        }
    }

    public static void Main()
    {
        // Using the property in the following manner
        // results in 201 copies of the array.
        // One copy is made each time the loop executes,
        // and one copy is made each time the condition is
        // tested.
        var t = new PC_DoNotReturnArrayFromProperty();

        for (int i = 0; i < t.Names.Length; i++)
        {
            if (t.Names[i] == ("SomeName"))
            {
                // Perform some operation.
            }
        }
    }

}
