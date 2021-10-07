using System;
using System.Text;

// This is a namespace
namespace RoslynCSSyntax
{
    // This is a class
    public class Class1
    {
        #region Constructors, Initialization, and Load

        // This is a Constructor
        public Class1()
        {
        }

        #endregion

        #region Enums

        // This is an enum
        public enum MyEnum
        {
            First_Element,
            Second_Element,
            Third_Element,
            Fourth_Element
        }

        #endregion

        #region Structures

        // This is a struct
        public struct MyStruct
        {
            private object _fieldMyObject;
            private string _fieldMyString;
            private long _fieldMyLong;
        }

        public struct YourStruct
        {
            private object _fieldYourObject;
            private string _fieldYourString;
            private long _fieldYourLong;
        }

        #endregion

        #region Fields and Properties


        private object _MyField;

        public string PropertyAutoString { get; set; }

        private string _propertyString;

        public string PropertyString
        {
            get => _propertyString;
            set => _propertyString = value;
        }

        #endregion

        #region Event Handlers


        #endregion

        #region Public Methods

        public void PublicMethodVoid()
        {
            int i = 0;
            int j;
            string s = "Hello";

            StringBuilder sb = new StringBuilder();
        }

        [Obsolete("Do not use")] // Attribute list
        public static /* modifiers */
        int // return type
        MyCoolMethod // Identifier
        <T> // Arity
        (T when, string where) // parameters
        where T : class // constraint clause
        {
            // body
            return 42;
        }

        #endregion

        #region Protected Methods

        protected void ProtectedMethodVoid()
        {

        }

        protected internal void ProtectedInternalMethodVoid()
        {

        }

        #endregion

        #region Private Methods


        // This is a method
        private bool PrivateMethodBool()
        {

            // This is a comment
            return false;
        }
        
        /// <summary>
        /// Does Something Cool
        /// </summary>
        /// <returns></returns>        
        public void LoggingMethod()
        {
#if LOGGING
            long startTicks = Log.PRESENTATION("Enter", Common.LOG_CATEGORY);
#endif


#if LOGGING
            Log.PRESENTATION($"End: filesToProcess.Count {filesToProcess.Count}", Common.LOG_CATEGORY, startTicks);
#endif
        }

        #endregion

    }
}
