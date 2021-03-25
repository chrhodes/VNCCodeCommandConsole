using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Prism.Events;

namespace VNCCodeCommandConsole.Core.Events
{
    // Put this in Core\Events
    public class ImportsStatementWalkerEvent : PubSubEvent { }
    // or
    //    public class ImportsStatementWalkerEvent : PubSubEvent<ImportsStatementWalkerEvent> { }
}
