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
                if (operationExpressions.Length > paramInfo.Length)
                    throw new ArgumentException(
                        $"提供了{operationExpressions.Length}个参数，但函数{method}仅需 要{paramInfo.Length}个参数。");
                for (int i = 0; i < paramInfo.Length; i++)
                {
                    if (i >= operationExpressions.Length ||
                        string.IsNullOrEmpty(operationExpressions[i]) && paramInfo[i].IsOptional)
                    {
                        arguments.Add(Type.Missing);
                    }
                    else if (paramInfo[i].ParameterType.IsEnum)
                    {
                        arguments.Add(Enum.Parse(paramInfo[i].ParameterType, operationExpressions[i], true));
                    }
                    else
                    {
                        arguments.Add(Convert.ChangeType(operationExpressions[i], paramInfo[i].ParameterType));
                    }
                }
                var c = (Closure) method.Invoke(this, arguments.ToArray());
                Debug.Assert(c != null);
                return c;
            }

            protected virtual void OnLeavingClosure()
            {
                
            }
        }

        /// <summary>
        /// 根。
        /// </summary>
        private class DocumentClosure : Closure
        {
            private float _SlideWidth;
            private float _SlideHeight;

            public Presentation Presentation { get; }

            public Application Application { get; }

            public float SlideWidth => _SlideWidth;

            public float SlideHeight => _SlideHeight;

            public string WorkPath { get; set; }

            public bool IsDebug { get; set; }

            private void UpdateCache()
            {
                _SlideWidth = Presentation.PageSetup.SlideWidth;
                _SlideHeight = Presentation.PageSetup.SlideHeight;
            }

            public void Initialize()
            {
                UpdateCache();
            }

            [ClosureOperation]
            public Closure Dir(string path)
            {
                WorkPath = path;
                return this;
            }

            [ClosureOperation]
            public Closure Debug(bool value)
            {
                IsDebug = value;
                return this;
            }

            [ClosureOperation]
            public Closure Page(string imagePath = "", string primaryText = "")
            {
                var closure = new PageClosure(this, Presentation.Slides.Add(
                    Presentation.Slides.Count + 1, PpSlideLayout.ppLayoutBlank));
                if (!string.IsNullOrEmpty(imagePath)) imagePath = Path.Combine(WorkPath, imagePath);
                closure.Initialize(imagePath, primaryText);
                return closure;
            }

            public DocumentClosure(Closure parent, Presentation presentation, Application application) : base(parent)
            {
                Presentation = presentation;
                Application = application;
                UpdateCache();
            }
        }

        private class AnimationInfo
        {
            public AnimationInfo(TimeSpan startAt, TimeSpan duration)
            {
                StartAt = startAt;
                Duration = duration;
            }

            public TimeSpan StartAt { get; set; }

            public TimeSpan Duration { get; set; }

            public TimeSpan EndAt => StartAt + Duration;

            public void ApplyTiming(Timing t)
            {
                if (t == null) throw new ArgumentNullException(nameof(t));
                t.TriggerDelayTime = (float) StartAt.TotalSeconds;
                t.Duration = (float)Duration.TotalSeconds;
            }

            public void ApplyTiming(IEnumerable<Timing> timings)
            {
                var lastEndAt = StartAt;
                foreach (var t in timings)
                {
                    t.TriggerDelayTime = (float)lastEndAt.TotalSeconds;
                    t.Duration = (float)Duration.TotalSeconds;
                    lastEndAt += Duration;
                }
                Duration = lastEndAt - StartAt;
            }
        }

        /// <summary>
        /// 一张图片的滚动方向（视图滚动方向）。
        /// </summary>
        public enum ImageScrollDirection
        {
            LeftUp,
            RightDown
        }

        private class PageClosure : Closure
        {
            const float MinImageScrollTime = 4; //Minimum

            public Slide Slide { get; }

            public Shape PrimaryImage { get; set; }

            public float ImageScrollTime { get; set; }

            public Shape PrimaryTextBox { get; set; }

            public float PagePersistTime { get; set; } = 1;

            private DocumentClosure Document => Ancestor<DocumentClosure>();

            private const int SubtitleFontSize = 24;
            private const int SecondarySubtitleFontSize = 20;

            private Effect primaryImageAnimation;
            private AnimationBehavior primaryImageSlideBehavior;

            private bool isPrimaryImageVertical;

            /// <summary>
            /// 由用户注册的所有动画。
            /// </summary>
            public List<AnimationInfo> Animations { get; } = new List<AnimationInfo>();

            public AnimationInfo AddAnimation(Shape shape, MsoAnimEffect effect,
                TimeSpan delay, TimeSpan duration, AnimationOptions options)
            {
                var e = Slide.TimeLine.MainSequence.AddEffect(shape, effect,
                    trigger: MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                e.Exit = (options & AnimationOptions.Exit) == AnimationOptions.Exit
                    ? MsoTriState.msoTrue
                    : MsoTriState.msoFalse;
                var ani = (options & AnimationOptions.AfterPrevious) == AnimationOptions.AfterPrevious
                    ? new AnimationInfo(Animations.LastOrDefault()?.EndAt ?? TimeSpan.Zero, duration)
                    : new AnimationInfo(Animations.LastOrDefault()?.StartAt ?? TimeSpan.Zero, duration);
                ani.ApplyTiming(e.Timing);
                if ((options & AnimationOptions.ByCharacter) == AnimationOptions.ByCharacter)
                {
                    e = Slide.TimeLine.MainSequence.ConvertToTextUnitEffect(e,
                        MsoAnimTextUnitEffect.msoAnimTextUnitEffectByCharacter);
                }
                if ((options & AnimationOptions.ByParagraph) == AnimationOptions.ByParagraph)
                {
                    const MsoAnimateByLevel BuildLevel = MsoAnimateByLevel.msoAnimateTextByAllLevels;
                    e = Slide.TimeLine.MainSequence.ConvertToBuildLevel(e, BuildLevel);
                    //此时 e 指向第一个段落的动画。
                    var paragrahAnimations = new List<Effect>();
                    for (int i = e.Index, j = Slide.TimeLine.MainSequence.Count; i <= j; i++)
                    {
                        var subEffect = Slide.TimeLine.MainSequence[i];
                        if (subEffect.Shape != shape
                            || subEffect.EffectInformation.BuildByLevelEffect != BuildLevel)
                            break;
                        paragrahAnimations.Add(subEffect);
                    }
                    //为每一个段落动画应用相同的触发模式。
                    ani.ApplyTiming(paragrahAnimations.Select(e1 => e1.Timing));
                }
                Animations.Add(ani);
                return ani;
            }

            public Shape CreateTextBox(string content, float y)
            {

                var textBox = Slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    0, y, Document.SlideWidth, 100);
                textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                textBox.TextFrame.WordWrap = MsoTriState.msoTrue;
                textBox.TextFrame.TextRange.Text = ParseTextExpression(content);
                textBox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                textBox.TextFrame.TextRange.Font.Size = SubtitleFontSize;
                textBox.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                textBox.TextFrame.TextRange.Font.Shadow = MsoTriState.msoTrue;
                textBox.TextFrame2.TextRange.Font.Line.Visible = MsoTriState.msoTrue;
                textBox.TextFrame2.TextRange.Font.Line.ForeColor.RGB = 0;
                textBox.TextFrame2.TextRange.Font.Line.Weight = 1;
                return textBox;
            }

            public void Initialize(string imagePath, string primaryText)
            {
                ImageScrollTime = MinImageScrollTime;
                var ppEntryEffect = RandomEntryEffect();
                Slide.SlideShowTransition.EntryEffect = ppEntryEffect;
                Slide.SlideShowTransition.Duration = 1;
                if (!string.IsNullOrEmpty(imagePath))
                {
                    PrimaryImage = Slide.Shapes.AddPicture(imagePath,
                        MsoTriState.msoTrue, MsoTriState.msoTrue, 0, 0);
                    primaryImageAnimation = Slide.TimeLine.MainSequence.AddEffect(PrimaryImage,
                        effectId: MsoAnimEffect.msoAnimEffectCustom,
                        trigger: MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    //宽 / 高
                    var sizeRatio = PrimaryImage.Width/PrimaryImage.Height;
                    var screenRatio = Document.SlideWidth/Document.SlideHeight;
                    // 图片比屏幕更高一些，纵向。
                    isPrimaryImageVertical = sizeRatio < screenRatio;
                    if (isPrimaryImageVertical)
                        PrimaryImage.Width = Document.SlideWidth;
                    else
                        PrimaryImage.Height = Document.SlideHeight;
                    primaryImageSlideBehavior = primaryImageAnimation.Behaviors.Add(MsoAnimType.msoAnimTypeMotion);
                    ImageDirection(isPrimaryImageVertical ? ImageScrollDirection.LeftUp : ImageScrollDirection.RightDown);
                    var scrollOffsetRelative = isPrimaryImageVertical
                        ? PrimaryImage.Height/Document.SlideHeight
                        : PrimaryImage.Width/Document.SlideWidth;
                    if (Math.Abs(scrollOffsetRelative) > 2)
                    {
                        // 如果超出屏幕的区域太长……
                        ImageScrollTime = MinImageScrollTime*Math.Abs(scrollOffsetRelative);
                    }
                    primaryImageAnimation.Timing.Duration = ImageScrollTime;
                    primaryImageAnimation.Timing.SmoothEnd = MsoTriState.msoCTrue;
                    if (Document.IsDebug)
                    {
                        var tb = Slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 50, 50);
                        tb.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                        tb.TextFrame.WordWrap = MsoTriState.msoFalse;
                        tb.TextFrame.TextRange.Font.Size = 9;
                        tb.TextFrame.TextRange.Text = imagePath;
                    }
                }
                if (!string.IsNullOrEmpty(primaryText))
                {
                    PrimaryTextBox = CreateTextBox(primaryText, 0);
                    PrimaryTextBox.TextFrame.TextRange.Font.Size = 24;
                    PrimaryTextBox.Top = Document.SlideHeight - PrimaryTextBox.Height;
                }
            }

            public Closure Persist(float persistTime)
            {
                PagePersistTime = persistTime;
                return this;
            }

            [ClosureOperation]
            public Closure ImageDirection(ImageScrollDirection direction)
            {
                if (isPrimaryImageVertical)
                {
                    // 纵向
                    PrimaryImage.Width = Document.SlideWidth;
                    var scrollOffsetRelative = (PrimaryImage.Height - Document.SlideHeight)/Document.SlideHeight;
                    if (direction == ImageScrollDirection.RightDown)
                        scrollOffsetRelative = -scrollOffsetRelative;
                    PrimaryImage.Top = direction == ImageScrollDirection.RightDown
                        ? 0
                        : Document.SlideHeight - PrimaryImage.Height;
                    primaryImageSlideBehavior.MotionEffect.Path =
                        $"M 0 0 L 0 {scrollOffsetRelative}";
                }
                else
                {
                    // 横向
                    PrimaryImage.Height = Document.SlideHeight;
                    var scrollOffsetRelative = (PrimaryImage.Width - Document.SlideWidth)/Document.SlideWidth;
                    if (direction == ImageScrollDirection.RightDown)
                        scrollOffsetRelative = -scrollOffsetRelative;
                    PrimaryImage.Left = direction == ImageScrollDirection.RightDown
                        ? 0
                        : Document.SlideWidth - PrimaryImage.Width;
                    primaryImageSlideBehavior.MotionEffect.Path =
                        $"M 0 0 L {scrollOffsetRelative} 0";
                }
                return this;
            }

            [ClosureOperation]
            public Closure Text(string text = "")
            {
                var closure = new TextClosure(this) {TextBox = CreateTextBox(text, 0)};
                //closure.EnterAnimation = Slide.TimeLine.MainSequence.AddEffect(closure.TextBox,
                //    effectId: MsoAnimEffect.msoAnimEffectFade,
                //    trigger: MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                //closure.EnterAnimation.Timing.Duration = TEXT_ADVANCE_TIME;
                return closure;
            }

            [ClosureOperation]
            public Closure Subtitle2(string text = "")
            {
                var closure = new TextClosure(this) { TextBox = CreateTextBox(text, 0) };
                closure.FontSize(SecondarySubtitleFontSize);
                closure.Bottom();
                return closure;
            }

            /// <summary>
            /// 插入BGM，注意，此功能仅用作打样。建议后期手动合成音轨。
            /// </summary>
            [ClosureOperation]
            public Closure Music(string path = "", int stopAfterSlides = 999)
            {
                var media = Slide.Shapes.AddMediaObject2(Path.Combine(Document.WorkPath, path));
                var effect = Slide.TimeLine.MainSequence.AddEffect(media, MsoAnimEffect.msoAnimEffectMediaPlay,
                    trigger: MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                effect.EffectInformation.PlaySettings.StopAfterSlides = stopAfterSlides;
                return this;
            }


            [ClosureOperation]
            public Closure Transition(PpEntryEffect effect = PpEntryEffect.ppEffectNone)
            {
                Slide.SlideShowTransition.EntryEffect = effect;
                return this;
            }

            protected override void OnLeavingClosure()
            {
                base.OnLeavingClosure();
                Slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
                var animationTime = Animations.Count > 0 ? Animations.Last().EndAt : TimeSpan.Zero;
                Slide.SlideShowTransition.AdvanceTime = Math.Max(ImageScrollTime, (float) animationTime.TotalSeconds) +
                                                        PagePersistTime;
            }

            public PageClosure(Closure parent, Slide slide) : base(parent)
            {
                Slide = slide;
            }
        }

        [Flags]
        public enum AnimationOptions
        {
            None = 0,
            Exit = 1,
            ByParagraph = 2,
            ByCharacter = 4,
            AfterPrevious = 8,
        };

        private class TextClosure : Closure
        {
            public const float TextAnimationDuration = 0.5f;

            public Shape TextBox { get; set; }

            public Effect EnterAnimation { get; set; }

            private DocumentClosure Document => Ancestor<DocumentClosure>();

            private PageClosure Page => Ancestor<PageClosure>();

            [ClosureOperation]
            public Closure Left(float value = 0)
            {
                TextBox.Left = value * Document.SlideWidth;
                return this;
            }

            [ClosureOperation]
            public Closure Top(float value = 0)
            {
                TextBox.Top = value * Document.SlideHeight;
                return this;
            }
            
            [ClosureOperation]
            public Closure Bottom(float value = 0)
            {
                TextBox.Top = Document.SlideHeight - (Page.PrimaryTextBox?.Height ?? 0)
                              - TextBox.Height + value*Document.SlideHeight;
                return this;
            }

            [ClosureOperation]
            public Closure VCenter(float offset = 0)
            {
                TextBox.Top = (Document.SlideHeight - TextBox.Height)/2 + offset*Document.SlideHeight;
                return this;
            }

            [ClosureOperation]
            public Closure Animation(MsoAnimEffect effect = MsoAnimEffect.msoAnimEffectFade,
                AnimationOptions options = AnimationOptions.None, float delay = 0)
            {
                Page.AddAnimation(TextBox, effect, TimeSpan.FromSeconds(delay), TimeSpan.FromSeconds(TextAnimationDuration), options);
                return this;
            }

            [ClosureOperation]
            public Closure FontSize(float size = 24)
            {
                TextBox.TextFrame.TextRange.Font.Size = size;
                return this;
            }

            [ClosureOperation]
            public Closure Paragraph(string text = "")
            {
                if (TextBox.TextFrame.HasText == MsoTriState.msoTrue)
                    TextBox.TextFrame.TextRange.InsertAfter("\r\n");
                TextBox.TextFrame.TextRange.InsertAfter(text);
                return this;
            }

            public TextClosure(Closure parent) : base(parent)
            {
            }
        }

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
                Parameters = parametersExpression.Split('\t');
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
