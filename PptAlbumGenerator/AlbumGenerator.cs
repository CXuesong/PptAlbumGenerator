using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Vbe.Interop;
using PptAlbumGenerator.Properties;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using ColorFormat = Microsoft.Office.Interop.PowerPoint.ColorFormat;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PptAlbumGenerator
{
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    internal sealed class ClosureOperationAttribute : Attribute
    {
    }

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

        private void ChangeTheme()
        {
            //TODO
            var path = @"C:\Program Files\Microsoft Office\Document Themes 16\Office Theme.thmx";
            var variantID = app.OpenThemeFile(path).ThemeVariants[3].Id;
            thisPresentation.ApplyTemplate2(path, variantID);
            thisPresentation.PageSetup.SlideSize = PpSlideSizeType.ppSlideSizeOnScreen;
        }

        private class Closure
        {
            public Closure(Closure parent)
            {
                Parent = parent;
            }

            public Closure Parent { get; }

            public void NotifyLeavingClosure()
            {
                OnLeavingClosure();
            }

            public T Ancestor<T>() where T : Closure
            {
                return Ancestor<T>(false);
            }

            public T Ancestor<T>(bool errorIfNotFound) where T : Closure
            {
                var c = this.Parent;
                while (c != null)
                {
                    var ancestor = c as T;
                    if (ancestor != null) return ancestor;
                    c = c.Parent;
                }
                if (errorIfNotFound)
                    throw new ArgumentException($"在当前作用域找不到{typeof (T)}类型的域。");
                return null;
            }

            public Closure InvokeOperation(string operationName, params string[] operationExpressions)
            {
                var method = GetType().GetMethod(operationName,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (method == null) throw new ArgumentException($"在{this}中找不到操作“{operationName}”。");
                var arguments = new List<object>();
                var paramInfo = method.GetParameters();
                for (int i = 0; i < operationExpressions.Length; i++)
                {
                    if (string.IsNullOrEmpty(operationExpressions[i]) && paramInfo[i].IsOptional)
                        arguments.Add(Type.Missing);
                    else
                        arguments.Add(Convert.ChangeType(operationExpressions[i], paramInfo[i].ParameterType));
                }
                return (Closure) method.Invoke(this, arguments.ToArray());
            }

            protected virtual void OnLeavingClosure()
            {
                
            }
        }

        private class DocumentClosure : Closure
        {
            private float _SlideWidth;
            private float _SlideHeight;

            public Presentation Presentation { get; }

            public Application Application { get; }

            public float SlideWidth => _SlideWidth;

            public float SlideHeight => _SlideHeight;

            private void UpdateCache()
            {
                _SlideWidth = Presentation.PageSetup.SlideWidth;
                _SlideHeight = Presentation.PageSetup.SlideHeight;
            }


            [ClosureOperation]
            public Closure Page(string imagePath, string primaryText)
            {
                var closure = new PageClosure(this, Presentation.Slides.Add(
                    Presentation.Slides.Count + 1, PpSlideLayout.ppLayoutBlank));

                return closure;
            }

            public DocumentClosure(Closure parent, Presentation presentation, Application application) : base(parent)
            {
                Presentation = presentation;
                Application = application;
                UpdateCache();
            }
        }

        private class PageClosure : Closure
        {
            public Slide Slide { get; }
            public Shape PrimaryImage { get; set; }
            public float ImageScrollTime { get; set; }
            public Shape PrimaryTextBox { get; set; }

            private DocumentClosure Document => Ancestor<DocumentClosure>();

            private PageClosure Page => Ancestor<PageClosure>();

            public Shape CreateTextBox(string content, float y)
            {

                var textBox = Page.Slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    0, y, Document.SlideWidth, 100);
                textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                textBox.TextFrame.WordWrap = MsoTriState.msoTrue;
                textBox.TextFrame.TextRange.Text = ParseTextExpression(content);
                textBox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                textBox.TextFrame.TextRange.Font.Size = 24;
                textBox.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                textBox.TextFrame.TextRange.Font.Shadow = MsoTriState.msoTrue;
                textBox.TextFrame2.TextRange.Font.Line.Visible = MsoTriState.msoTrue;
                textBox.TextFrame2.TextRange.Font.Line.ForeColor.RGB = 0;
                textBox.TextFrame2.TextRange.Font.Line.Weight = 1;
                return textBox;
            }

            [ClosureOperation]
            public Closure Text(string text)
            {
                var closure = new TextClosure(this) {TextBox = CreateTextBox(text, 0)};

                return closure;
            }

            public PageClosure(Closure parent, Slide slide) : base(parent)
            {
                Slide = slide;
            }
        }

        private class TextClosure : Closure
        {
            public Shape TextBox { get; set; }


            public TextClosure(Closure parent) : base(parent)
            {
            }
        }

        private class ScopeStatus : ICloneable
        {
            public event EventHandler<EventArgs> ExitingScope;

            public int Indension { get; set; }
            /// <summary>当前的工作路径。</summary>
            public string WorkPath { get; set; }

            public Slide Slide { get; set; }

            public float ImageScrollTime { get; set; }

            public Shape TitleTextBox { get; set; }

            /// <summary>
            /// 当前页面最后一个使用 TEXT 指令插入的文本框。
            /// </summary>
            public List<Shape> TextBoxes { get; } = new List<Shape>();

            public string ExtendPath(string path)
            {
                return Path.Combine(WorkPath, path);
            }

            public ScopeStatus Clone()
            {
                var inst = (ScopeStatus)MemberwiseClone();
                inst.ExitingScope = null;
                return inst;
            }

            object ICloneable.Clone()
            {
                return Clone();
            }

            internal virtual void OnExitingScope()
            {
                ExitingScope?.Invoke(this, EventArgs.Empty);
            }
        }

        private void ThrowInvalidIndension()
        {
            throw new FormatException($"无效的缩进。");
        }

        private ScopeStatus CurrentScope => scopeStack.Peek();

        private void EnterScope(int indension)
        {
            var inst = scopeStack.Peek().Clone();
            scopeStack.Push(inst);
            inst.Indension = indension;
        }

        private void ExitScope()
        {
            scopeStack.Pop().OnExitingScope();
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
                Parameters = parametersExpression.Split('\t');
            }
        }

        private readonly Stack<ScopeStatus> scopeStack = new Stack<ScopeStatus>();

        private static readonly Regex lineMatcher =
            new Regex(@"^(?<Indension>\s*)(?<Command>\S*)((\s)(?<Params>.*))?$");

        private float slideWidth, slideHeight;

        private Effect MakePathAnimation(Shape shape, MsoAnimTriggerType trigger)
        {
            var effect = (Effect) app.Run("PAG_MakePathAnimation", shape);
            effect.Timing.TriggerType = trigger;
            return effect;
        }

        private static string ParseTextExpression(string expr)
        {
            if (string.IsNullOrEmpty(expr)) return "";
            return expr.Replace('|', '\n');
        }

        public Shape CreateSubtitle(string content, float y)
        {

            var textBox = CurrentScope.Slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                0, y, thisPresentation.PageSetup.SlideWidth, 100);
            textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            textBox.TextFrame.WordWrap = MsoTriState.msoTrue;
            textBox.TextFrame.TextRange.Text = ParseTextExpression(content);
            textBox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            textBox.TextFrame.TextRange.Font.Size = 24;
            textBox.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            textBox.TextFrame.TextRange.Font.Shadow = MsoTriState.msoTrue;
            textBox.TextFrame2.TextRange.Font.Line.Visible = MsoTriState.msoTrue;
            textBox.TextFrame2.TextRange.Font.Line.ForeColor.RGB = 0;
            textBox.TextFrame2.TextRange.Font.Line.Weight = 1;
            return textBox;
        }

        private void ParseLine(string line)
        {
            const float PAGE_PERSIST_TIME = 4;      //Minimum
            const float TEXT_ADVANCE_TIME = 0.5f;
            const float TEXT_PERSIST_TIME = 2f;
            var match = lineMatcher.Match(line);
            if (!match.Success) throw new FormatException($"无法识别的行：{line}");
            var instruction = new Instruction(match.Groups["Indension"].Value.Length, match.Groups["Command"].Value, match.Groups["Params"].Value);
            while (instruction.Indension <= CurrentScope.Indension)
                ExitScope();
            switch (instruction.Command)
            {
                case "DIR":
                    CurrentScope.WorkPath = instruction.ParametersExpression;
                    break;
                case "PAGE":
                    {
                        EnterScope(instruction.Indension);
                        CurrentScope.ExitingScope += (sender, e) =>
                        {
                            //重新计算幻灯片持续时间。
                            var scope = (ScopeStatus) sender;
                            scope.Slide.SlideShowTransition.AdvanceTime = Math.Max(
                                scope.ImageScrollTime,
                                scope.TextBoxes.Count*(TEXT_ADVANCE_TIME + TEXT_PERSIST_TIME) + TEXT_ADVANCE_TIME);
                        };
                        CurrentScope.TitleTextBox = null;
                        CurrentScope.TextBoxes.Clear();
                        CurrentScope.Slide = thisPresentation.Slides.Add(
                            thisPresentation.Slides.Count + 1,
                            PpSlideLayout.ppLayoutBlank);
                        CurrentScope.ImageScrollTime = PAGE_PERSIST_TIME;
                        CurrentScope.Slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
                        var ppEntryEffect = RandomEntryEffect();
                        CurrentScope.Slide.SlideShowTransition.EntryEffect = ppEntryEffect;
                        CurrentScope.Slide.SlideShowTransition.Duration = 1;
                        var picturePath = instruction.ParameterAt(0);
                        if (!string.IsNullOrEmpty(picturePath))
                        {
                            picturePath = CurrentScope.ExtendPath(picturePath);
                            //解决换页时黑屏的问题。
                            //var pictureOverlay = CurrentScope.Slide.Shapes.AddPicture(picturePath,
                            //    MsoTriState.msoTrue, MsoTriState.msoTrue, 0, 0);
                            var picture = CurrentScope.Slide.Shapes.AddPicture(picturePath,
                                MsoTriState.msoTrue, MsoTriState.msoTrue, 0, 0);
                            var effect = CurrentScope.Slide.TimeLine.MainSequence.AddEffect(picture,
                                effectId: MsoAnimEffect.msoAnimEffectCustom,
                                trigger: MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            //宽 / 高
                            var sizeRatio = picture.Width/picture.Height;
                            float scrollOffsetRelative;
                            if (sizeRatio > 1)
                            {
                                // 横向
                                picture.Height = thisPresentation.PageSetup.SlideHeight;
                                scrollOffsetRelative = -(picture.Width - slideWidth)/slideWidth;
                                //var effect = MakePathAnimation(picture, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                effect.Behaviors.Add(MsoAnimType.msoAnimTypeMotion).MotionEffect.Path =
                                    $"M 0 0 L {scrollOffsetRelative} 0";
                            }
                            else
                            {
                                // 纵向
                                picture.Width = thisPresentation.PageSetup.SlideWidth;
                                picture.Top = slideHeight - picture.Height;
                                scrollOffsetRelative = (picture.Height - slideHeight)/slideHeight;
                                //var effect = MakePathAnimation(picture, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                effect.Behaviors.Add(MsoAnimType.msoAnimTypeMotion).MotionEffect.Path =
                                    $"M 0 0 L 0 {scrollOffsetRelative}";
                            }
                            if (Math.Abs(scrollOffsetRelative) > 2)
                            {
                                // 如果宽高比过于极端
                                CurrentScope.ImageScrollTime =
                                    PAGE_PERSIST_TIME*(float) Math.Ceiling(Math.Abs(scrollOffsetRelative));
                            }
                            effect.Timing.Duration = CurrentScope.ImageScrollTime;
                            effect.Timing.SmoothEnd = MsoTriState.msoCTrue;
                            picture.AnimationSettings.AdvanceMode = PpAdvanceMode.ppAdvanceOnTime;
                            picture.ZOrder(MsoZOrderCmd.msoBringToFront);
                            //CurrentScope.Slide.Background.Fill.UserTextured(pic);
                            //CurrentScope.Slide.Background.Fill.TextureAlignment = MsoTextureAlignment.msoTextureCenter;
                        }
                        var textBox = CreateSubtitle(ParseTextExpression(instruction.ParameterAt(1)), 0);
                        textBox.TextFrame.TextRange.Font.Size = 24;
                        textBox.Top = slideHeight - textBox.Height;
                        CurrentScope.TitleTextBox = textBox;
                    }
                    break;
                case "TEXT":
                    if (instruction.HasParameter(1))
                    {
                        var textBox = CreateSubtitle(ParseTextExpression(instruction.ParameterAt(1)), 0);
                        textBox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                        textBox.TextFrame.TextRange.Font.Size = 21;
                        textBox.Top = CurrentScope.TitleTextBox.Top - textBox.Height -
                                      CurrentScope.TitleTextBox.Height*0.2f;
                        var effect = instruction.HasParameter(0)
                            ? (MsoAnimEffect)
                                typeof (MsoAnimEffect).GetField("msoAnimEffect" + instruction.ParameterAt(0),
                                    BindingFlags.Static | BindingFlags.Public | BindingFlags.IgnoreCase).GetValue(null)
                            : MsoAnimEffect.msoAnimEffectFade;
                        var enterAnimation = CurrentScope.Slide.TimeLine.MainSequence.AddEffect(textBox,
                            effect,
                            trigger: MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        enterAnimation.Timing.Duration = TEXT_ADVANCE_TIME;
                        if (CurrentScope.TextBoxes.Count > 0)
                        {
                            var exitAnimation = CurrentScope.Slide.TimeLine.MainSequence.AddEffect(
                                CurrentScope.TextBoxes.Last(),
                                MsoAnimEffect.msoAnimEffectFade,
                                trigger: MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            exitAnimation.Exit = MsoTriState.msoTrue;
                            exitAnimation.Timing.TriggerDelayTime= enterAnimation.Timing.TriggerDelayTime =
                                CurrentScope.TextBoxes.Count*(TEXT_ADVANCE_TIME + TEXT_PERSIST_TIME);
                            exitAnimation.Timing.Duration = TEXT_ADVANCE_TIME / 2;
                        }
                        CurrentScope.TextBoxes.Add(textBox);
                    }
                    break;
                default:
                    throw new InvalidOperationException($"无法识别的指令：{instruction.Command}。");
            }
        }

        public void Generate()
        {
            scopeStack.Clear();
            scopeStack.Push(new ScopeStatus() { Indension = -1 });
            if (app == null) app = new Application();
            thisPresentation = app.Presentations.Add(MsoTriState.msoFalse);
            ChangeTheme();
            slideWidth = thisPresentation.PageSetup.SlideWidth;
            slideHeight = thisPresentation.PageSetup.SlideHeight;
            var module = thisPresentation.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            module.CodeModule.AddFromString(Resources.VBAModule);
            while (true)
            {
                var thisLine = ScriptReader.ReadLine();
                if (thisLine == null) break;
                Console.WriteLine(thisLine);
                if (!string.IsNullOrWhiteSpace(thisLine) && thisLine[0] != '#')
                    ParseLine(thisLine);
            }
            while (CurrentScope.Indension >= 0)
                ExitScope();
            app.Run("PAG_PostProcess", thisPresentation);
            thisPresentation.NewWindow();
            //thisPresentation.VBProject.VBComponents.Remove(module); 
            //thisPresentation.Close();
        }
    }
}
