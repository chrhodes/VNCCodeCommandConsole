using Prism.Events;

using static VNC.CodeAnalysis.CommandTypes;

namespace VNCCodeCommandConsole.Core.Events
{
    // Put this in Core\Events
    public class InvokeCSSyntaxWalkerEvent : PubSubEvent<SearchTreeCommand> { }
    // or
    //    public class ImportsStatementWalkerEvent : PubSubEvent<ImportsStatementWalkerEvent> { }
}
