using System;

namespace foo
{ 
	public class ClassPublic
	{
		public ClassPublic()
		{
		}
		public void Method1()
        {
            
        }
    }

	private class ClassPrivate
	{
		public ClassPrivate()
		{
		}

		private void Method1()
		{

		}
	}

	protected class ClassProtected
	{
		public ClassProtected()
		{
		}

		protected void Method1()
		{

		}
	}

    protected internal class ClassProtectedInternal
    {
        public ClassProtectedInternal()
        {
        }

        protected void Method1()
        {

        }
    }

    internal class ClassInternal
	{
		public ClassInternal()
		{
		}

		internal void Method1()
		{

		}
	}
}

