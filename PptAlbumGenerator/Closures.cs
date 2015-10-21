using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PptAlbumGenerator
{
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    internal sealed class ClosureOperationAttribute : Attribute
    {
    }

    internal class Closure
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
                throw new ArgumentException($"�ڵ�ǰ�������Ҳ���{typeof (T)}���͵���");
            return null;
        }

        public Closure InvokeOperation(string operationName, params string[] operationExpressions)
        {
            var method = GetType().GetMethod(operationName,
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
            if (method == null) throw new ArgumentException($"��{this}���Ҳ���������{operationName}����");
            var arguments = new List<object>();
            var paramInfo = method.GetParameters();
            if (operationExpressions.Length > paramInfo.Length)
                throw new ArgumentException(
                    $"�ṩ��{operationExpressions.Length}��������������{method}���� Ҫ{paramInfo.Length}��������");
            for (int i = 0; i < paramInfo.Length; i++)
            {
                if (i >= operationExpressions.Length ||
                    String.IsNullOrEmpty(operationExpressions[i]) && paramInfo[i].IsOptional)
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


    internal class AnimationInfo
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
            t.TriggerDelayTime = (float)StartAt.TotalSeconds;
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

    [Flags]
    public enum AnimationOptions
    {
        None = 0,
        Exit = 1,
        ByParagraph = 2,
        ByCharacter = 4,
        AfterPrevious = 8,
    };

    /// <summary>
    /// һ��ͼƬ�Ĺ���������ͼ�������򣩡�
    /// </summary>
    public enum ImageScrollDirection
    {
        LeftUp,
        RightDown
    }

    internal class TextClosure : Closure
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
                          - TextBox.Height + value * Document.SlideHeight;
            return this;
        }

        [ClosureOperation]
        public Closure VCenter(float offset = 0)
        {
            TextBox.Top = (Document.SlideHeight - TextBox.Height) / 2 + offset * Document.SlideHeight;
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

    internal class PageClosure : Closure
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
        /// ���û�ע������ж�����
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
            var ani =
                new AnimationInfo((((options & AnimationOptions.AfterPrevious) == AnimationOptions.AfterPrevious
                    ? Animations.LastOrDefault()?.EndAt
                    : Animations.LastOrDefault()?.StartAt) ?? TimeSpan.Zero) + delay,
                    duration);
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
                //��ʱ e ָ���һ������Ķ�����
                var paragrahAnimations = new List<Effect>();
                for (int i = e.Index, j = Slide.TimeLine.MainSequence.Count; i <= j; i++)
                {
                    var subEffect = Slide.TimeLine.MainSequence[i];
                    if (subEffect.Shape != shape
                        || subEffect.EffectInformation.BuildByLevelEffect != BuildLevel)
                        break;
                    paragrahAnimations.Add(subEffect);
                }
                //Ϊÿһ�����䶯��Ӧ����ͬ�Ĵ���ģʽ��
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
            textBox.TextFrame.TextRange.Text = content;
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
            var ppEntryEffect = AlbumGenerator.RandomEntryEffect();
            Slide.SlideShowTransition.EntryEffect = ppEntryEffect;
            Slide.SlideShowTransition.Duration = 1;
            if (!String.IsNullOrEmpty(imagePath))
            {
                PrimaryImage = Slide.Shapes.AddPicture(imagePath,
                    MsoTriState.msoTrue, MsoTriState.msoTrue, 0, 0);
                primaryImageAnimation = Slide.TimeLine.MainSequence.AddEffect(PrimaryImage,
                    effectId: MsoAnimEffect.msoAnimEffectCustom,
                    trigger: MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                //�� / ��
                var sizeRatio = PrimaryImage.Width / PrimaryImage.Height;
                var screenRatio = Document.SlideWidth / Document.SlideHeight;
                // ͼƬ����Ļ����һЩ������
                isPrimaryImageVertical = sizeRatio < screenRatio;
                if (isPrimaryImageVertical)
                    PrimaryImage.Width = Document.SlideWidth;
                else
                    PrimaryImage.Height = Document.SlideHeight;
                primaryImageSlideBehavior = primaryImageAnimation.Behaviors.Add(MsoAnimType.msoAnimTypeMotion);
                ImageDirection(isPrimaryImageVertical ? ImageScrollDirection.LeftUp : ImageScrollDirection.RightDown);
                var scrollOffsetRelative = isPrimaryImageVertical
                    ? PrimaryImage.Height / Document.SlideHeight
                    : PrimaryImage.Width / Document.SlideWidth;
                if (Math.Abs(scrollOffsetRelative) > 2)
                {
                    // ���������Ļ������̫������
                    ImageScrollTime = MinImageScrollTime * Math.Abs(scrollOffsetRelative);
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
            if (!String.IsNullOrEmpty(primaryText))
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
                // ����
                PrimaryImage.Width = Document.SlideWidth;
                var scrollOffsetRelative = (PrimaryImage.Height - Document.SlideHeight) / Document.SlideHeight;
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
                // ����
                PrimaryImage.Height = Document.SlideHeight;
                var scrollOffsetRelative = (PrimaryImage.Width - Document.SlideWidth) / Document.SlideWidth;
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
            var closure = new TextClosure(this) { TextBox = CreateTextBox(text, 0) };
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
        /// ����BGM��ע�⣬�˹��ܽ�������������������ֶ��ϳ����졣
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
            Slide.SlideShowTransition.AdvanceTime = Math.Max(ImageScrollTime, (float)animationTime.TotalSeconds) +
                                                    PagePersistTime;
        }

        public PageClosure(Closure parent, Slide slide) : base(parent)
        {
            Slide = slide;
        }
    }

    /// <summary>
    /// ����
    /// </summary>
    internal class DocumentClosure : Closure
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
            if (!String.IsNullOrEmpty(imagePath)) imagePath = Path.Combine(WorkPath, imagePath);
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
}