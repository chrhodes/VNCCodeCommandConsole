using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace VNC.CodeAnalysis.Helpers
{
    public static class CS
    {
        public static List<string> GetMethodNames(string sourceCode)
        {
            List<string> methodNames = new List<string>();

            SyntaxTree tree = CSharpSyntaxTree.ParseText(sourceCode);

            VNC.CodeAnalysis.SyntaxWalkers.CS.MethodDeclarationNames walker = null;
            walker = new VNC.CodeAnalysis.SyntaxWalkers.CS.MethodDeclarationNames();
            walker.MethodNames = methodNames;
            walker.Visit(tree.GetRoot());

            return methodNames;
        }

        public static string GetContainingMethodName(CSharpSyntaxNode node)
        {
            string methodName = "none";

            var a5 = node.Ancestors()
                .Where(x => x.IsKind(SyntaxKind.MethodDeclaration))
                .Cast<MethodDeclarationSyntax>().ToList();

            if (a5.Count > 0)
            {
                methodName = node.Ancestors()
                    .Where(x => x.IsKind(SyntaxKind.MethodDeclaration))
                    .Cast<MethodDeclarationSyntax>().First().Identifier.ToString();
            }

            return methodName;
        }

        public static string GetContainingMethodBlock(CSharpSyntaxNode node)
        {
            string methodBlock = "<METHODBLOCK>";

            var a5 = node.Ancestors()
                .Where(x => x.IsKind(SyntaxKind.MethodDeclaration))
                .Cast<MethodDeclarationSyntax>().ToList();

            if (a5.Count > 0)
            {
                methodBlock = node.Ancestors()
                    .Where(x => x.IsKind(SyntaxKind.MethodDeclaration))
                    .Cast<MethodDeclarationSyntax>().FirstOrDefault().ToString();
            }

            return methodBlock;
        }

        public static string GetContainingContext(CSharpSyntaxNode node, CodeAnalysisOptions displayInfo)
        {
            string ancestorContext = GetAncestorContext(node, displayInfo);

            string classModuleContext = GetClassModuleContext(node, displayInfo);

            string methodContext = GetMethodContext(node, displayInfo);

            string sourceContext = GetSourceContext(node, displayInfo);

            return ancestorContext + classModuleContext + methodContext + sourceContext;
        }

        private static string GetSourceContext(CSharpSyntaxNode node, CodeAnalysisOptions displayInfo)
        {
            string sourceContext = "";

            if (displayInfo.DisplaySourceLocation)
            {
                var location = node.GetLocation();
                var sourceSpan = location.SourceSpan;
                var lineSpan = location.GetLineSpan();

                //// NB.  Lines start at 0.  Add one so when we look in Visual Studio it makes sense.

                //var startLine = location.GetLineSpan().StartLinePosition.Line + 1;
                //var endLine = location.GetLineSpan().EndLinePosition.Line + 1;

                sourceContext = string.Format("SourceSpan: {0} LineSpan: {1}",
                    sourceSpan.ToString(),
                    lineSpan.ToString());
            }

            return sourceContext;
        }

        private static string GetMethodContext(CSharpSyntaxNode node, CodeAnalysisOptions displayInfo)
        {
            string methodContext = "";

            if (displayInfo.DisplayMethodName)
            {
                methodContext += string.Format(" Method:({0, -35})", Helpers.CS.GetContainingMethodName(node));
            }
            else if (displayInfo.DisplayContainingMethodBlock)
            {
                methodContext += string.Format(" MethodBlock:({0})", Helpers.CS.GetContainingMethodBlock(node));
            }

            return methodContext;
        }

        private static string GetClassModuleContext(CSharpSyntaxNode node, CodeAnalysisOptions displayInfo)
        {
            string classModuleContext = "";

            if (displayInfo.DisplayClassOrModuleName)
            {
                var inClassBlock = node.Ancestors()
                    .Where(x => x.IsKind(SyntaxKind.ClassDeclaration))
                    .Cast<ClassDeclarationSyntax>().ToList();

                //var inModuleBlock = node.Ancestors()
                //    .Where(x => x.IsKind(SyntaxKind.ModuleBlock))
                //    .Cast<ModuleBlockSyntax>().ToList();

                string typeName = "unknown";
                string className = "unknown";
                string moduleName = "unknown";

                if (inClassBlock.Count > 0)
                {
                    typeName = "Class";

                    className = node.Ancestors()
                        .Where(x => x.IsKind(SyntaxKind.ClassDeclaration))
                        .Cast<ClassDeclarationSyntax>().First().Identifier.ToString();
                }

                //if (inModuleBlock.Count > 0)
                //{
                //    typeName = "Module";

                //    moduleName = node.Ancestors()
                //        .Where(x => x.IsKind(SyntaxKind.ModuleBlock))
                //        .Cast<ModuleBlockSyntax>().First().ModuleStatement.Identifier.ToString();
                //}

                classModuleContext = String.Format("{0, 8}{1,6}:({2,-25})",
                    classModuleContext,
                    typeName,
                    typeName == "Class" ? className : moduleName);
            }

            return classModuleContext;
        }

        private static string GetAncestorContext(CSharpSyntaxNode node, CodeAnalysisOptions displayInfo)
        {
            string ancestorContext = "";

            if (displayInfo.DisplayContainingBlock)
            {
                ancestorContext += GetContainingBlock(node).Kind().ToString();
            }

            if (displayInfo.InTryBlock)
            {
                var inTryBlock = node.Ancestors()
                    .Where(x => x.IsKind(SyntaxKind.TryStatement))
                    .Cast<TryStatementSyntax>().ToList();

                if (inTryBlock.Count > 0)
                {
                    ancestorContext += "T ";
                }
            }

            if (displayInfo.InWhileBlock)
            {
                var inDoWhileBlock = node.Ancestors()
                    .Where(x => x.IsKind(SyntaxKind.WhileStatement))
                    .Cast<WhileStatementSyntax>().ToList();

                if (inDoWhileBlock.Count > 0)
                {
                    ancestorContext += "W ";
                }
            }

            if (displayInfo.InForBlock)
            {
                var inForBlock = node.Ancestors()
                    .Where(x => x.IsKind(SyntaxKind.ForStatement))
                    .Cast<ForStatementSyntax>().ToList();

                if (inForBlock.Count > 0)
                {
                    ancestorContext += "F ";
                }
            }

            if (displayInfo.InIfBlock)
            {
                var inMultiLineIfBlock = node.Ancestors()
                    .Where(x => x.IsKind(SyntaxKind.IfStatement))
                    .Cast<IfStatementSyntax>().ToList();

                if (inMultiLineIfBlock.Count > 0)
                {
                    ancestorContext += "I ";
                }
            }

            if (ancestorContext.Length > 0)
            {
                ancestorContext = string.Format("{0,8}", ancestorContext);
            }

            return ancestorContext;
        }

        public static StringBuilder InvokeVNCSyntaxWalker(
            SyntaxWalkers.CS.VNCCSSyntaxWalkerBase walker,
            SearchTreeCommandConfiguration commandConfiguration)
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", CodeAnalysis.Common.LOG_CATEGORY);

            StringBuilder results = new StringBuilder();

            //walker.Messages = commandConfiguration.Results;
            walker.Messages = results;

            // Setting the TargetPattern will call InitializeRegEx()
            walker.TargetPattern = commandConfiguration.UseRegEx ? commandConfiguration.RegEx : ".*";

            walker._configurationOptions = commandConfiguration.ConfigurationOptions;

            walker.Matches = commandConfiguration.Matches;
            walker.CRCMatchesToString = commandConfiguration.CRCMatchesToString;
            walker.CRCMatchesToFullString = commandConfiguration.CRCMatchesToFullString;

            walker.Visit(commandConfiguration.SyntaxTree.GetRoot());

            if (results.Length > 0)
            {
                if (commandConfiguration.ConfigurationOptions.DisplayCRC32)
                {
                    results.AppendFormat("CRC32Node:            {0}\n", walker.CRC32Node);
                    results.AppendFormat("CRC32NodeKind:        {0}\n", walker.CRC32NodeKind);
                    //results.AppendFormat("CRC32Token:           {0}\n", walker.CRC32Token);
                    //results.AppendFormat("CRC32Trivia:          {0}\n", walker.CRC32Trivia);
                    //results.AppendFormat("CRC32StructuredTrivia:{0}\n", walker.CRC32StructuredTrivia);
                }

                commandConfiguration.Results.AppendLine(results.ToString());
            }

            Log.DOMAINSERVICES("Exit", CodeAnalysis.Common.LOG_CATEGORY, startTicks);

            return commandConfiguration.Results;
        }

        public static StringBuilder InvokeVNCSyntaxRewriter(
            SyntaxRewriters.CS.VNCCSSyntaxRewriterBase rewriter,
            RewriteTreeCommandConfiguration commandConfiguration)
        {
            StringBuilder results = new StringBuilder();

            //walker.Messages = commandConfiguration.Results;
            rewriter.Messages = results;

            rewriter.TargetPattern = commandConfiguration.UseRegEx ? commandConfiguration.TargetPattern : ".*";
            rewriter._configurationOptions = commandConfiguration.ConfigurationOptions;

            rewriter._targetPatternRegEx = Common.InitializeRegEx(rewriter.TargetPattern, rewriter.Messages, RegexOptions.IgnoreCase);
            //walker.InitializeRegEx();

            rewriter.Replacements = commandConfiguration.Replacements;
            //rewriter.CRCMatchesToString = commandConfiguration.CRCMatchesToString;
            //rewriter.CRCMatchesToFullString = commandConfiguration.CRCMatchesToFullString;

            rewriter.Visit(commandConfiguration.SyntaxTree.GetRoot());

            if (results.Length > 0)
            {
                //if (commandConfiguration.ConfigurationOptions.ShowBlockCRC)
                //{
                //    results.AppendFormat("CRC32Node:            {0}\n", walker.CRC32Node);
                //    results.AppendFormat("CRC32Token:           {0}\n", walker.CRC32Token);
                //    results.AppendFormat("CRC32Trivia:          {0}\n", walker.CRC32Trivia);
                //    results.AppendFormat("CRC32StructuredTrivia:{0}\n", walker.CRC32StructuredTrivia);
                //}

                commandConfiguration.Results.AppendLine(results.ToString());
            }

            return commandConfiguration.Results;
        }

        // TODO(crhodes)
        // Not sure we need the typed versions.  Maybe just the VisualBasicSyntaxNode version

        public static Boolean IsOnLineByItself(InvocationExpressionSyntax node)
        {
            Boolean result = false;

            // Seems like Trivia is Trivia  No notion of with our without :)

            string existingLeadingTrivia = node.GetLeadingTrivia().ToString();
            //string existingLeadingTriviaFull = node.GetLeadingTrivia().ToFullString();

            string existingTrailingTrivia = node.GetTrailingTrivia().ToString();
            //string existingTrailingTriviaFull = node.GetTrailingTrivia().ToFullString();

            if (String.IsNullOrWhiteSpace(existingLeadingTrivia) && String.IsNullOrWhiteSpace(existingTrailingTrivia))
            {
                result = true;
            }
            else
            {
                // TODO(crhodes)
                // There are lots of different types of Trivia.  Verify this is enough.

                var triviaList = node.GetLeadingTrivia().ToSyntaxTriviaList();

                Boolean leadingTriviaResult = false;

                foreach (SyntaxTrivia syntaxTrivia in node.GetLeadingTrivia().ToSyntaxTriviaList())
                {
                    var triviaKind = syntaxTrivia.Kind();

                    if (triviaKind == SyntaxKind.EndOfLineTrivia
                        || triviaKind == SyntaxKind.WhitespaceTrivia
                        || triviaKind == SyntaxKind.MultiLineCommentTrivia
                        || triviaKind == SyntaxKind.SingleLineCommentTrivia)
                    {
                        leadingTriviaResult = true;
                    }
                }

                Boolean trailingTriviaResult = false;

                foreach (SyntaxTrivia syntaxTrivia in node.GetTrailingTrivia().ToSyntaxTriviaList())
                {
                    var triviaKind = syntaxTrivia.Kind();

                    if (triviaKind == SyntaxKind.EndOfLineTrivia
                        || triviaKind == SyntaxKind.WhitespaceTrivia
                        || triviaKind == SyntaxKind.MultiLineCommentTrivia
                        || triviaKind == SyntaxKind.SingleLineCommentTrivia)
                    {
                        trailingTriviaResult = true;
                    }
                }

                result = leadingTriviaResult & trailingTriviaResult;
            }

            return result;
        }

        public static Boolean IsOnLineByItself(ExpressionStatementSyntax node)
        {
            Boolean result = false;

            // Seems like Trivia is Trivia  No notion of with our without :)

            string existingLeadingTrivia = node.GetLeadingTrivia().ToString();
            //string existingLeadingTriviaFull = node.GetLeadingTrivia().ToFullString();

            string existingTrailingTrivia = node.GetTrailingTrivia().ToString();
            //string existingTrailingTriviaFull = node.GetTrailingTrivia().ToFullString();

            if (String.IsNullOrWhiteSpace(existingLeadingTrivia) && String.IsNullOrWhiteSpace(existingTrailingTrivia))
            {
                result = true;
            }
            else
            {
                var triviaList = node.GetLeadingTrivia().ToSyntaxTriviaList();

                Boolean leadingTriviaResult = false;

                foreach (SyntaxTrivia syntaxTrivia in node.GetLeadingTrivia().ToSyntaxTriviaList())
                {
                    var triviaKind = syntaxTrivia.Kind();

                    if (triviaKind == SyntaxKind.EndOfLineTrivia
                        || triviaKind == SyntaxKind.WhitespaceTrivia
                        || triviaKind == SyntaxKind.MultiLineCommentTrivia
                        || triviaKind == SyntaxKind.SingleLineCommentTrivia)
                    {
                        leadingTriviaResult = true;
                    }
                }

                Boolean trailingTriviaResult = false;

                foreach (SyntaxTrivia syntaxTrivia in node.GetTrailingTrivia().ToSyntaxTriviaList())
                {
                    var triviaKind = syntaxTrivia.Kind();

                    if (triviaKind == SyntaxKind.EndOfLineTrivia
                        || triviaKind == SyntaxKind.WhitespaceTrivia
                        || triviaKind == SyntaxKind.MultiLineCommentTrivia
                        || triviaKind == SyntaxKind.SingleLineCommentTrivia)
                    {
                        trailingTriviaResult = true;
                    }
                }

                result = leadingTriviaResult & trailingTriviaResult;
            }

            return result;
        }

        public static Boolean IsOnLineByItself(CSharpSyntaxNode node)
        {
            Boolean result = false;

            // Seems like Trivia is Trivia  No notion of with our without :)

            string existingLeadingTrivia = node.GetLeadingTrivia().ToString();
            //string existingLeadingTriviaFull = node.GetLeadingTrivia().ToFullString();

            string existingTrailingTrivia = node.GetTrailingTrivia().ToString();
            //string existingTrailingTriviaFull = node.GetTrailingTrivia().ToFullString();

            if (String.IsNullOrWhiteSpace(existingLeadingTrivia) && String.IsNullOrWhiteSpace(existingTrailingTrivia))
            {
                result = true;
            }
            else
            {
                var triviaList = node.GetLeadingTrivia().ToSyntaxTriviaList();

                Boolean leadingTriviaResult = false;

                foreach (SyntaxTrivia syntaxTrivia in node.GetLeadingTrivia().ToSyntaxTriviaList())
                {
                    var triviaKind = syntaxTrivia.Kind();

                    if (triviaKind == SyntaxKind.EndOfLineTrivia
                        || triviaKind == SyntaxKind.WhitespaceTrivia
                        || triviaKind == SyntaxKind.MultiLineCommentTrivia
                        || triviaKind == SyntaxKind.SingleLineCommentTrivia)
                    {
                        leadingTriviaResult = true;
                    }
                }

                Boolean trailingTriviaResult = false;

                foreach (SyntaxTrivia syntaxTrivia in node.GetTrailingTrivia().ToSyntaxTriviaList())
                {
                    var triviaKind = syntaxTrivia.Kind();

                    if (triviaKind == SyntaxKind.EndOfLineTrivia
                        || triviaKind == SyntaxKind.WhitespaceTrivia
                        || triviaKind == SyntaxKind.MultiLineCommentTrivia
                        || triviaKind == SyntaxKind.SingleLineCommentTrivia)
                    {
                        trailingTriviaResult = true;
                    }
                }

                result = leadingTriviaResult & trailingTriviaResult;
            }

            return result;
        }

        /// <summary>
        /// Takes a (multi-line) comment and returns a leading whitespace aware comment.
        /// </summary>
        /// <param name="comment"></param>
        /// <param name="leadingWhiteSpace"></param>
        /// <returns></returns>
        public static string MultiLineComment(string comment, string leadingWhiteSpace)
        {
            StringBuilder result = new StringBuilder();

            var lines = Regex.Split(comment, "\r\n|\r|\n");

            foreach (string line in lines)
            {
                var newLine = string.Format("'{0}", line);

                result.AppendLine(string.Format("'{0}\n{1}", line, leadingWhiteSpace));
            }

            return result.ToString();
        }

        public static string BlockComment(string block)
        {
            StringBuilder result = new StringBuilder();

            var lines = Regex.Split(block, "\r\n|\r|\n");

            foreach (string line in lines)
            {
                var newLine = string.Format("'{0}", line);

                result.AppendLine(newLine);
            }

            return result.ToString();
        }

        static CSharpSyntaxNode GetContainingBlock(CSharpSyntaxNode node)
        {
            //var block = node.Parent as VisualBasicSyntaxNode;
            //var blockKind = block.Kind();
            var blockKindText = node.Parent.Kind().ToString();

            // TODO(crhodes)
            // This may not work in CSharp as Block is not in all the typically Namespace, Class, Module, Method, like VB

            if (blockKindText.Contains("Block"))
            {
                return (CSharpSyntaxNode)node.Parent;
            }

            if (blockKindText == "CompilationUnit")
            {
                return node;
            }

            return GetContainingBlock(node.Parent as CSharpSyntaxNode);
        }
    }
}
