// From Source Code Analysis with Roslyn -

using System;
using System.Threading;

public class PC_DoNotUseThreadAbortOrThreadSuspend
{
    public static void Main()
    {
        Thread newThread =
        new Thread(new ThreadStart(TestMethod));
        newThread.Start();
        Thread.Sleep(1000);
        // Abort newThread.
        Console.WriteLine("Main aborting new thread.");
        newThread.Abort("Information from Main.");
        // Wait for the thread to terminate.
        newThread.Join();
        Console.WriteLine(@"New thread terminated – Main exiting.");
    }

    private static void TestMethod()
    {

    }
}