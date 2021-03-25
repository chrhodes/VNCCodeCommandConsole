using System.Collections.Generic;

using Microsoft.CodeAnalysis;

using VNC.Core.Mvvm;

using VNCCA = VNC.CodeAnalysis;

namespace VNCCodeCommandConsole.Presentation.ModelWrappers
{
    public class CodeAnalysisOptionsWrapper : ModelWrapper<VNCCA.CodeAnalysisOptions>
    {
        public CodeAnalysisOptionsWrapper() { }
        public CodeAnalysisOptionsWrapper(VNCCA.CodeAnalysisOptions model) : base(model)
        {
        }

        // TODO(crhodes)
        // Wrap each property from the passed in model.

        public SyntaxWalkerDepth SyntaxWalkerDepth { get { return GetValue<SyntaxWalkerDepth>(); } set { SetValue(value); } }

        #region Output Options

        public VNCCA.SyntaxNode.AdditionalNodes AdditionalNodeAnalysis { get { return GetValue<VNCCA.SyntaxNode.AdditionalNodes>(); } set { SetValue(value); } }

        public bool DisplayNodeKind { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool DisplayNodeValue { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool DisplayFormattedOutput { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool DisplayNodeParent { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool DisplayStatementBlock { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool IncludeStatementBlockInCRC { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool DisplaySourceLocation { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool DisplayCRC32 { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool ShowAnalysisCRC { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool ReplaceCRLFInNodeName { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool DisplayClassOrModuleName { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool DisplayMethodName { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool DisplayContainingMethodBlock { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool DisplayContainingBlock { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool InTryBlock { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool InWhileBlock { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool InForBlock { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool InIfBlock { get { return GetValue<bool>(); } set { SetValue(value); } }


        public bool AllTypes { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool Boolean { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool Byte { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool DataTable { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool Date { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool DateTime { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool Int16 { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool Int32 { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool Int64 { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool Integer { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool Long { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool Single { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool String { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool OtherTypes { get { return GetValue<bool>(); } set { SetValue(value); } }

        #endregion

        #region Rewriter Options

        public bool AddFileSuffix { get { return GetValue<bool>(); } set { SetValue(value); } }

        public string FileSuffix { get { return GetValue<string>(); } set { SetValue(value); } }

        #endregion

    }
}
