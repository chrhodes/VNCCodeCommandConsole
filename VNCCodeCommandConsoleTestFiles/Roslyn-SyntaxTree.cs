using System;

// This is a namespace
namespace Namespace1
{
    // This is a class
    public class Class1
    {
#region Constructor

        public Class1()
        {
        }
        
#endregion

        // This is a method
        private bool PrivateMethodBool()
        {

            // This is a comment
        }
        
        /// <summary>
        /// Does Something Cool
        /// </summary>
        /// <returns></returns>        
        public void PublicMethodVoid()
        {
#ifdef LOGGING
            long startTicks = Log.PRESENTATION("Enter", Common.LOG_CATEGORY);
#endif


#ifdef LOGGING
            Log.PRESENTATION($"End: filesToProcess.Count {filesToProcess.Count}", Common.LOG_CATEGORY, startTicks);
#endif
        }

        [Obsolete("Do not use")] // Attribute list
        public static /* modifiers */
        int // return type
        MyCoolMethod // Identifier
        <T> // Arity
        (T when, string where) // parameters
        where T:class // constraint clause
        {
            // body
            return 42;
        }
    }
}
