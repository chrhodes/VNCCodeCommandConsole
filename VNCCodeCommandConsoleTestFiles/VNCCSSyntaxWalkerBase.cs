using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

using Crc32C;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace VNC.CodeAnalysis.SyntaxWalkers.CS
{

    public class VNCCSSyntaxWalkerBase : CSharpSyntaxWalker
    {
        public StringBuilder Messages;
        public StringBuilder WalkerNode = new StringBuilder();
        public StringBuilder WalkerToken = new StringBuilder();
        public StringBuilder WalkerTrivia = new StringBuilder();
        public StringBuilder WalkerStructuredTrivia = new StringBuilder();

        public CodeAnalysisOptions _configurationOptions = new CodeAnalysisOptions();

        private string _targetPattern;
        internal Regex _targetPatternRegEx;

        internal VNC.CodeAnalysis.SyntaxNode.FieldDeclarationLocation _declarationLocation;

        public Dictionary<string, Int32> Matches;
        public Dictionary<string, Int32> CRCMatchesToString;
        public Dictionary<string, Int32> CRCMatchesToFullString;

        public string CRC32Node;
        public string CRC32NodeKind;
        public string CRC32Token;
        public string CRC32Trivia;
        public string CRC32StructuredTrivia;

        public ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();

        public string TargetPattern
        {
            get => _targetPattern;

            set
            {
                _targetPattern = value;
                _targetPatternRegEx = Helpers.Common.InitializeRegEx(_targetPattern, Messages, RegexOptions.IgnoreCase);
            }
        }

        public VNCCSSyntaxWalkerBase(SyntaxWalkerDepth depth = SyntaxWalkerDepth.Node) : base(depth)
        {

        }

        public string GetNodeContext(CSharpSyntaxNode node)
        {
            string messageContext = "";

            // NOTE(crhodes)
            // CSharp does not have concept of Module.  This Config option should be ok, however.

            if (_configurationOptions.DisplayClassOrModuleName)
            {
                messageContext = Helpers.CS.GetContainingContext(node, _configurationOptions);
            }

            if (_configurationOptions.DisplayMethodName)
            {
                messageContext += string.Format(" Method:({0, -35})", Helpers.CS.GetContainingMethodName(node));
            }

            return messageContext;
        }

        // NOTE(crhodes)
        // CSharp does not have Module Blocks.
        // In CSharp this is really Namespace, Class, Method, Block

        public enum BlockType : Int16
        {
            None = 0,
            NamespaceBlock = 1,
            ClassBlock = 2,
            ModuleBlock = 3,
            MethodBlock = 4
        }

        public void RecordMatch(CSharpSyntaxNode node, BlockType blockType)
        {
            string nodeValue = "";

            // Produce a useful string to use for the Matches dictionary
            // Don't want to use a long string

            switch (blockType)
            {
                case BlockType.None:
                    nodeValue = node.ToString();
                    break;

                case BlockType.NamespaceBlock:
                    nodeValue = ((NamespaceDeclarationSyntax)node).Name.ToString();
                    break;

                case BlockType.ClassBlock:
                    nodeValue = ((ClassDeclarationSyntax)node).Identifier.ToString();
                    break;

                //case BlockType.ModuleBlock:
                //    nodeValue = ((ModuleBlockSyntax)node).ModuleStatement.Identifier.ToString();
                //    break;

                case BlockType.MethodBlock:
                    nodeValue = ((MethodDeclarationSyntax)node).Identifier.ToString();
                    break;
            }

            if (_configurationOptions.ReplaceCRLFInNodeName)
            {
                nodeValue = nodeValue.Replace("\r\n", " ");
            }

            if (_configurationOptions.DisplayCRC32)
            {
                byte[] nodeToStringBytes = asciiEncoding.GetBytes(node.ToString());
                byte[] nodeToFullStringBytes = asciiEncoding.GetBytes(node.ToFullString());
                string toStringCRC = Crc32CAlgorithm.Compute(nodeToStringBytes).ToString();
                string toFullStringCRC = Crc32CAlgorithm.Compute(nodeToFullStringBytes).ToString();

                Messages.AppendLine(String.Format("{0, -35} CRC32:({1,10}) ({2,10})",
                    nodeValue,
                    toStringCRC,
                    toFullStringCRC));

                string toStringKey = nodeValue + ":" + toStringCRC;
                string toFullStringKey = nodeValue + ":" + toFullStringCRC;

                // The Node

                if (CRCMatchesToString.ContainsKey(toStringKey))
                {
                    CRCMatchesToString[toStringKey] += 1;
                }
                else
                {
                    CRCMatchesToString.Add(toStringKey, 1);
                }

                // The Node and Trivia

                if (CRCMatchesToFullString.ContainsKey(toFullStringKey))
                {
                    CRCMatchesToFullString[toFullStringKey] += 1;
                }
                else
                {
                    CRCMatchesToFullString.Add(toFullStringKey, 1);
                }
            }
            else
            {
                Messages.AppendLine(String.Format("{0}",
                    nodeValue));
            }

            if (Matches.ContainsKey(nodeValue))
            {
                Matches[nodeValue] += 1;
            }
            else
            {
                Matches.Add(nodeValue, 1);
            }
        }

        public void RecordMatchAndContext(CSharpSyntaxNode node, BlockType blockType)
        {
            string nodeValue = "";
            string nodeKey = "";

            // Produce a useful string to use for the Matches dictionary
            // Don't want to use a long string
            //
            // Also handle display of full block for some SyntaxNodes

            switch (blockType)
            {
                case BlockType.None:
                    nodeValue = node.ToString();
                    nodeKey = nodeValue;
                    break;

                case BlockType.NamespaceBlock:
                    nodeValue = ((NamespaceDeclarationSyntax)node).Name.ToString();
                    // TODO(crhodes)
                    // May want to remove .Name
                    nodeKey = ((NamespaceDeclarationSyntax)node).Name.ToString();
                    break;

                case BlockType.ClassBlock:
                    nodeValue = ((ClassDeclarationSyntax)node).Identifier.ToString();
                    // TODO(crhodes)
                    // May want to remove .Identifier
                    nodeKey = ((ClassDeclarationSyntax)node).Identifier.ToString(); ;
                    break;

                //case BlockType.ModuleBlock:
                //    nodeValue = ((ModuleBlockSyntax)node).ToString();
                //    nodeKey = ((ModuleBlockSyntax)node).ModuleStatement.ToString();
                //    break;

                case BlockType.MethodBlock:
                    if (_configurationOptions.DisplayStatementBlock)
                    {
                        nodeValue = ((MethodDeclarationSyntax)node).ToString();
                    }
                    else
                    {
                        nodeValue = "<METHODBLOCK>";
                    }

                    //nodeValue = ((MethodBlockSyntax)node).ToString();
                    nodeKey = ((MethodDeclarationSyntax)node).ToString(); ;
                    break;
            }

            if (_configurationOptions.ReplaceCRLFInNodeName)
            {
                // TODO(crhodes)
                // This may not work if we want to see the block

                nodeKey = nodeKey.Replace("\r\n", " ");
            }

            if (_configurationOptions.DisplayCRC32)
            {
                byte[] nodeToStringBytes = asciiEncoding.GetBytes(node.ToString());
                byte[] nodeToFullStringBytes = asciiEncoding.GetBytes(node.ToFullString());
                string toStringCRC = Crc32CAlgorithm.Compute(nodeToStringBytes).ToString();
                string toFullStringCRC = Crc32CAlgorithm.Compute(nodeToFullStringBytes).ToString();

                // TODO(crhodes)
                // Decide if more useful to have CRC first or last.

                //Messages.AppendLine(String.Format("{0} {1,-35} CRC32:({2,10}) ({3,10})",
                Messages.AppendLine(String.Format("CRC32:({2,10}) ({3,10}) {0} {1,-35}",
                    Helpers.CS.GetContainingContext(node, _configurationOptions),
                    nodeValue,
                    toStringCRC,
                    toFullStringCRC));

                //string toStringKey = string.Format("{0}:({1,10})", nodeValue, toStringCRC);
                //string toFullStringKey = string.Format("{0}:({1,10})", nodeValue, toFullStringCRC);

                string toStringKey = string.Format("({0,10}):{1}", toStringCRC, nodeKey);
                string toFullStringKey = string.Format("({0,10}):{1}", toFullStringCRC, nodeKey);


                // The Node

                if (CRCMatchesToString.ContainsKey(toStringKey))
                {
                    CRCMatchesToString[toStringKey] += 1;
                }
                else
                {
                    CRCMatchesToString.Add(toStringKey, 1);
                }

                // The Node and Trivia

                if (CRCMatchesToFullString.ContainsKey(toFullStringKey))
                {
                    CRCMatchesToFullString[toFullStringKey] += 1;
                }
                else
                {
                    CRCMatchesToFullString.Add(toFullStringKey, 1);
                }
            }
            else
            {
                Messages.AppendLine(String.Format("{0} {1}",
                    Helpers.CS.GetContainingContext(node, _configurationOptions),
                    nodeValue));
            }

            // TODO(crhodes)
            // May want to just use the CRC stuff

            if (Matches.ContainsKey(nodeKey))
            {
                Matches[nodeKey] += 1;
            }
            else
            {
                Matches.Add(nodeKey, 1);
            }

            if (_configurationOptions.DisplayNodeParent)
            {
                Messages.AppendLine(String.Format("NodeParent:>{0}<", node.Parent.ToString()));
            }

            // This was first done in MethodBlock.  Now incorporating it generally.

            // Everything below here is an attempt to determine if a method 
            // is Syntactically the same as another method.
            //
            // Some challenges.
            // If only nodes kinds are evaluated things will show as the same even if different names or trivia
            //
            // If NodeValues are considered, at least two problems occur.
            // 1. If we evaluate the body of the method, any trivia in the body is evaluated.
            // 2. If we we don't evaluate the body, we have to iterate across the descendants

            //var walkerNode = new VNC.CodeAnalysis.SyntaxWalkers.VB.VisitAll(SyntaxWalkerDepth.Node);
            //var walkerToken = new VNC.CodeAnalysis.SyntaxWalkers.VB.VisitAll(SyntaxWalkerDepth.Token);
            //var walkerTrivia = new VNC.CodeAnalysis.SyntaxWalkers.VB.VisitAll(SyntaxWalkerDepth.Trivia);
            //var walkerStructuredTrivia = new VNC.CodeAnalysis.SyntaxWalkers.VB.VisitAll(SyntaxWalkerDepth.StructuredTrivia);

            // Get the contained MethodStatement

            IEnumerable<Microsoft.CodeAnalysis.SyntaxNode> analysisNodes = null;
            IEnumerable<Microsoft.CodeAnalysis.SyntaxToken> analysisTokens = null;
            IEnumerable<Microsoft.CodeAnalysis.SyntaxTrivia> analysisTrivia = null;
            IEnumerable<Microsoft.CodeAnalysis.SyntaxNodeOrToken> analysisNodesOrTokens = null;

            ChildSyntaxList childAnalysisNodes;

            switch (_configurationOptions.AdditionalNodeAnalysis)
            {
                case CodeAnalysis.SyntaxNode.AdditionalNodes.Ancestors:
                    analysisNodes = node.Ancestors();
                    break;

                case CodeAnalysis.SyntaxNode.AdditionalNodes.AncestorsAndSelf:
                    analysisNodes = node.AncestorsAndSelf();
                    break;

                case CodeAnalysis.SyntaxNode.AdditionalNodes.ChildNodes:
                    analysisNodes = node.ChildNodes();
                    break;

                case CodeAnalysis.SyntaxNode.AdditionalNodes.ChildNodesAndTokens:
                    childAnalysisNodes = node.ChildNodesAndTokens();
                    break;

                case CodeAnalysis.SyntaxNode.AdditionalNodes.ChildTokens:
                    analysisTokens = node.ChildTokens();
                    break;

                case CodeAnalysis.SyntaxNode.AdditionalNodes.DescendantNodes:
                    analysisNodes = node.DescendantNodes();
                    break;

                case CodeAnalysis.SyntaxNode.AdditionalNodes.DescendantNodesAndSelf:
                    analysisNodes = node.DescendantNodesAndSelf();
                    break;

                case CodeAnalysis.SyntaxNode.AdditionalNodes.DescendantNodesAndTokens:
                    analysisNodesOrTokens = node.DescendantNodesAndTokens();
                    break;

                case CodeAnalysis.SyntaxNode.AdditionalNodes.DescendantNodesAndTokensAndSelf:
                    analysisNodesOrTokens = node.DescendantNodesAndTokensAndSelf();
                    break;

                case CodeAnalysis.SyntaxNode.AdditionalNodes.DescendantTokens:
                    analysisTokens = node.DescendantTokens();
                    break;

                case CodeAnalysis.SyntaxNode.AdditionalNodes.DescendantTrivia:
                    analysisTrivia = node.DescendantTrivia();
                    break;

                case CodeAnalysis.SyntaxNode.AdditionalNodes.None:
                    break;

                default:
                    // ??
                    break;
            }

            StringBuilder analysisNodeKind = new StringBuilder();
            StringBuilder analysisNodeValue = new StringBuilder();

            bool displayAnalysisOutput = _configurationOptions.DisplayNodeKind || _configurationOptions.DisplayNodeValue;

            // TODO(crhodes)
            // See if we can using some clever casting to handle this more generically
            //
            // analysisNodes
            // analysisTokens
            // analysisTrivia
            // analysisNodesOrTokens

            if (analysisNodes != null)
            {
                if (displayAnalysisOutput) { Messages.AppendLine(">>> AnalysisNodes >>>"); }

                foreach (var item in analysisNodes)
                {
                    var isBlock = item.Kind().ToString().Contains("Block");

                    analysisNodeKind.AppendLine(string.Format("{0}", item.Kind().ToString()));

                    // Blocks bring along everything including trivia which messes up the CRC.  Just use "BLOCK"

                    analysisNodeValue.AppendLine(string.Format("{0}", isBlock ? "BLOCK" : item.ToString()));

                    if (displayAnalysisOutput)
                    {
                        Messages.AppendLine(string.Format("Node:{0}:>{1}<",
                            _configurationOptions.DisplayNodeKind ? item.Kind().ToString() : "",
                            _configurationOptions.DisplayNodeValue ? isBlock ? "BLOCK" : item.ToString() : ""));
                    }
                }

                if (displayAnalysisOutput) { Messages.AppendLine("<<< AnalysisNodes <<<"); }
            }

            if (analysisTokens != null)
            {
                if (displayAnalysisOutput) { Messages.AppendLine(">>> AnalysisTokens >>>"); }

                foreach (var item in analysisTokens)
                {
                    var isBlock = item.Kind().ToString().Contains("Block");

                    analysisNodeKind.AppendLine(string.Format("{0}", item.Kind().ToString()));

                    // Blocks bring along everything including trivia which messes up the CRC.  Just use "BLOCK"

                    analysisNodeValue.AppendLine(string.Format("{0}", isBlock ? "BLOCK" : item.ToString()));

                    if (displayAnalysisOutput)
                    {
                        Messages.AppendLine(string.Format("Node:{0}:>{1}<",
                            _configurationOptions.DisplayNodeKind ? item.Kind().ToString() : "",
                            _configurationOptions.DisplayNodeValue ? isBlock ? "BLOCK" : item.ToString() : ""));
                    }
                }

                if (displayAnalysisOutput) { Messages.AppendLine("<<< AnalysisTokens <<<"); }
            }

            if (analysisTrivia != null)
            {
                if (displayAnalysisOutput) { Messages.AppendLine(">>> AnalysisTrivia >>>"); }

                foreach (var item in analysisTrivia)
                {
                    var isBlock = item.Kind().ToString().Contains("Block");

                    analysisNodeKind.AppendLine(string.Format("{0}", item.Kind().ToString()));

                    // Blocks bring along everything including trivia which messes up the CRC.  Just use "BLOCK"

                    analysisNodeValue.AppendLine(string.Format("{0}", isBlock ? "BLOCK" : item.ToString()));

                    if (displayAnalysisOutput)
                    {
                        Messages.AppendLine(string.Format("Node:{0}:>{1}<",
                            _configurationOptions.DisplayNodeKind ? item.Kind().ToString() : "",
                            _configurationOptions.DisplayNodeValue ? isBlock ? "BLOCK" : item.ToString() : ""));
                    }
                }

                if (displayAnalysisOutput) { Messages.AppendLine("<<< AnalysisTrivia <<<"); }
            }

            if (analysisNodesOrTokens != null)
            {
                if (displayAnalysisOutput) { Messages.AppendLine(">>> AnalysisNodesOrTokens >>>"); }

                foreach (var item in analysisNodesOrTokens)
                {
                    var isBlock = item.Kind().ToString().Contains("Block");

                    analysisNodeKind.AppendLine(string.Format("{0}", item.Kind().ToString()));

                    // Blocks bring along everything including trivia which messes up the CRC.  Just use "BLOCK"

                    analysisNodeValue.AppendLine(string.Format("{0}", isBlock ? "BLOCK" : item.ToString()));

                    if (displayAnalysisOutput)
                    {
                        Messages.AppendLine(string.Format("Node:{0}:>{1}<",
                            _configurationOptions.DisplayNodeKind ? item.Kind().ToString() : "",
                            _configurationOptions.DisplayNodeValue ? isBlock ? "BLOCK" : item.ToString() : ""));
                    }
                }

                if (displayAnalysisOutput) { Messages.AppendLine("<<< AnalysisNodesOrTokens <<<"); }
            }

            byte[] analysisToBytes = asciiEncoding.GetBytes(analysisNodeKind.ToString());
            CRC32Node = Crc32CAlgorithm.Compute(analysisToBytes).ToString();

            analysisToBytes = asciiEncoding.GetBytes(analysisNodeValue.ToString());
            CRC32NodeKind = Crc32CAlgorithm.Compute(analysisToBytes).ToString();
        }

        /// <summary>
        /// Used by MethodStatement so we can pass in nodeValue
        /// </summary>
        /// <param name="node"></param>
        /// <param name="nodeValue"></param>
        public void RecordMatchAndContext(CSharpSyntaxNode node, string nodeValue)
        {
            if (_configurationOptions.DisplayCRC32)
            {
                byte[] nodeToStringBytes = asciiEncoding.GetBytes(node.ToString());
                byte[] nodeToFullStringBytes = asciiEncoding.GetBytes(node.ToFullString());
                string toStringCRC = Crc32CAlgorithm.Compute(nodeToStringBytes).ToString();
                string toFullStringCRC = Crc32CAlgorithm.Compute(nodeToFullStringBytes).ToString();

                // TODO(crhodes)
                // Decide if more useful to have CRC first or last.

                //Messages.AppendLine(String.Format("{0} {1,-35} CRC32:({2,10}) ({3,10})",
                Messages.AppendLine(String.Format("CRC32:({2,10}) ({3,10}) {0} {1,-35}",
                    Helpers.CS.GetContainingContext(node, _configurationOptions),
                    nodeValue,
                    toStringCRC,
                    toFullStringCRC));

                //string toStringKey = string.Format("{0}:({1,10})", nodeValue, toStringCRC);
                //string toFullStringKey = string.Format("{0}:({1,10})", nodeValue, toFullStringCRC);

                string toStringKey = string.Format("({0,10}):{1}", toStringCRC, nodeValue);
                string toFullStringKey = string.Format("({0,10}):{1}", toFullStringCRC, nodeValue);

                // The Node

                if (CRCMatchesToString.ContainsKey(toStringKey))
                {
                    CRCMatchesToString[toStringKey] += 1;
                }
                else
                {
                    CRCMatchesToString.Add(toStringKey, 1);
                }

                // The Node and Trivia

                if (CRCMatchesToFullString.ContainsKey(toFullStringKey))
                {
                    CRCMatchesToFullString[toFullStringKey] += 1;
                }
                else
                {
                    CRCMatchesToFullString.Add(toFullStringKey, 1);
                }
            }
            else
            {
                Messages.AppendLine(String.Format("{0} {1}",
                    Helpers.CS.GetContainingContext(node, _configurationOptions),
                    nodeValue));
            }

            // TODO(crhodes)
            // May want to just use the CRC stuff

            if (Matches.ContainsKey(nodeValue))
            {
                Matches[nodeValue] += 1;
            }
            else
            {
                Matches.Add(nodeValue, 1);
            }
        }
    }
}
