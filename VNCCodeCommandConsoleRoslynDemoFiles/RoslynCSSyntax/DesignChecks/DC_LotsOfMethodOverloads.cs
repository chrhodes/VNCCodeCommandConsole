// From Source Code Analysis with Roslyn -

using System;
using System.Reflection.Metadata;

public class DC_LotsOfMethodOverloads
{

        public Document createDocument(String name)
        {
            // just calls another method with default value
            // of its parameter
            return createDocument(name, -1);
        }

        public Document createDocument(String name, int minPagesCount)
        {
            // just calls another method with default value of its parameter
            return createDocument(name, minPagesCount, false);
        }

        public Document createDocument(String name, int minPagesCount, Boolean firstPageBlank)
        {
            // just calls another method with default value of its parameter
            return createDocument(name, minPagesCount, false, "");
        }

        public Document createDocument(String name, int minPagesCount, Boolean firstPageBlank, String title)
        {
            // here the real work gets done
            return new Document();
        }
  }
