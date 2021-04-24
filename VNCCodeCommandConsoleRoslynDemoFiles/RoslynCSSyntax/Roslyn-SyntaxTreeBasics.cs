using System;

namespace Namespace1
{
    public class Class1
    {
        public Class1()
        {
        }

        private bool PrivateMethodBool()
        {
            return false;
        }

        private string PrivateMethodString()
        {
            return "hello";
        }

        private void PrivateMethodVoid()
        {

        }

        internal void InternalMethodVoid()
        {

        }

        protected internal void ProtectedInternalMethodVoid()
        {

        }

        protected void ProtectedMethodVoid()
        {

        }
        public void PublicMethodVoid()
        {

        }

        private class Class1a
        {

        }

        private class Class1b
        {
            private class Class1b1
            {

            }
        }
    }

    namespace NestedNamespace
    {

        public class ClassOne
        {
            
            public ClassOne()
            {
                void Method1()
                {
                    
                }

                void Method2()
                {
                    
                }
            }

            private class ClassOneA
            {

            }

            private class ClassOneB
            {
                private class ClassOneBOne
                {

                }
            }
        }
    }
}
