
using Prism.Events;

namespace VNCCodeCommandConsole.Core.Events
{
    // Put this in Core\Events
    public class SyntaxWalkerResultEvent : PubSubEvent<string> { }
    // or
    //    public class ImportsStatementWalkerEvent : PubSubEvent<TYPE> { }
}
