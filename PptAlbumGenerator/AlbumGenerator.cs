using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Vbe.Interop;
using PptAlbumGenerator.Properties;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace PptAlbumGenerator
{
    public partial class AlbumGenerator
    {
        public AlbumGenerator(TextReader scriptReader)
        {
            if (scriptReader == null) throw new ArgumentNullException(nameof(scriptReader));
            ScriptReader = scriptReader;
        }

        public TextReader ScriptReader { get; }

        private Application app;
        private Presentation thisPresentation;

        private void ThrowInvalidIndension()
        {
            throw new FormatException($"无效的缩进。");
        }

        private Closure CurrentClosure;

        private void EnterClosure(Closure newClosure)
        {
            if (newClosure != CurrentClosure)
            {
                CurrentClosure = newClosure;
            }
        }

        private void ExitClosure()
        {
            CurrentClosure.NotifyLeavingClosure();
            CurrentClosure = CurrentClosure.Parent;
        }

        public class Instruction
        {
            public int Indension { get; }

            public string Command { get; }

            public string ParametersExpression { get; }

            public string[] Parameters { get; }

            public bool HasParameter(int index)
            {
                return !string.IsNullOrEmpty(ParameterAt(index));
            }

            public string ParameterAt(int index)
            {
                if (index >= Parameters.Length) return null;
                return Parameters[index];
            }

            public Instruction(int indension, string command, string parametersExpression)
            {
                Indension = indension;
                Command = command;
                ParametersExpression = parametersExpression;
                Parameters = parametersExpression.Split('\t').Select(ParseTextExpression).ToArray();
            }
        }

        private readonly Stack<int> indensionStack = new Stack<int>();

        private static readonly Regex lineMatcher =
            new Regex(@"^(?<Indension>\s*)(?<Command>\S*)((\s)(?<Params>.*))?$");

        private static string ParseTextExpression(string expr)
        {
            if (string.IsNullOrEmpty(expr)) return "";
            return expr.Replace('|', '\n');
        }

        private void ParseLine(string line)
        {
            var match = lineMatcher.Match(line);
            if (!match.Success) throw new FormatException($"无法识别的行：{line}");
            var instruction = new Instruction(match.Groups["Indension"].Value.Length, match.Groups["Command"].Value, match.Groups["Params"].Value);
            while (instruction.Indension <= indensionStack.Peek())
            {
                indensionStack.Pop();
                ExitClosure();
            }
            var newClosure = CurrentClosure.InvokeOperation(instruction.Command, instruction.Parameters);
            if (newClosure != CurrentClosure)
            {
                indensionStack.Push(instruction.Indension);
                Debug.Assert(newClosure.Parent == CurrentClosure);
                CurrentClosure = newClosure;
            }
        }
        private void ChangeTheme()
        {
            //TODO
            var path = @"C:\Program Files\Microsoft Office\Document Themes 16\Office Theme.thmx";
            var variantID = app.OpenThemeFile(path).ThemeVariants[3].Id;
            thisPresentation.ApplyTemplate2(path, variantID);
            thisPresentation.PageSetup.SlideSize = PpSlideSizeType.ppSlideSizeOnScreen;
        }

        public void Generate()
        {
            indensionStack.Clear();
            indensionStack.Push(-1);
            if (app == null) app = new Application();
            thisPresentation = app.Presentations.Add(MsoTriState.msoFalse);
            CurrentClosure = new DocumentClosure(null, thisPresentation, app);
            ChangeTheme();
            ((DocumentClosure) CurrentClosure).Initialize();
            var module = thisPresentation.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            module.CodeModule.AddFromString(Resources.VBAModule);
            while (true)
            {
                var thisLine = ScriptReader.ReadLine();
                if (thisLine == null) break;
                Console.WriteLine(thisLine);
                if (!string.IsNullOrWhiteSpace(thisLine) && thisLine.TrimStart()[0] != '#')
                    ParseLine(thisLine);
            }
            while (CurrentClosure != null) ExitClosure();
            app.Run("PAG_PostProcess", thisPresentation);
            var wnd = thisPresentation.NewWindow();
            thisPresentation.VBProject.VBComponents.Remove(module); 
            wnd.Activate();
            //thisPresentation.Close();
        }
    }
}
