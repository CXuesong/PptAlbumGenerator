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

        public SlideTransitionPoolClosure SlideTransitionPool { get; }

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
            WorkPath = Path.Combine(WorkPath, path);
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

        [ClosureOperation]
        public Closure Transitions()
        {
            return SlideTransitionPool;
        }

        public DocumentClosure(Closure parent, Presentation presentation, Application application) : base(parent)
        {
            Presentation = presentation;
            Application = application;
            SlideTransitionPool = new SlideTransitionPoolClosure(this);
            UpdateCache();
        }
    }

    internal enum PrimaryImageAnimation
    {
        None = 0,
        Fit,
        /// <summary>
        /// 视图从右向左，或从下向上。
        /// </summary>
        ScrollNear,
        /// <summary>
        /// 视图从左向右，或从上向下。
        /// </summary>
        ScrollFar,
        Expand,
        Shrink,
        ExpandOrShink,
    }

    internal class PageClosure : Closure
    {
        const float MinImageAnimationDuration = 4; //Minimum

        public Slide Slide { get; }

        public Shape PrimaryImage { get; set; }

        public float PrimaryImageAnimationDuration { get; set; }

        public Shape PrimaryTextBox { get; set; }

        public float PagePersistTime { get; set; } = 1;

        private DocumentClosure Document => Ancestor<DocumentClosure>();

        private const int SubtitleFontSize = 24;
        private const int SecondarySubtitleFontSize = 20;

        private Effect primaryImageAnimation;

        private bool isPrimaryImageVertical;

        /// <summary>
        /// 由用户注册的所有动画。
        /// </summary>
        public List<AnimationInfo> Animations { get; } = new List<AnimationInfo>();

        public void Initialize(string imagePath, string primaryText)
        {
            PrimaryImageAnimationDuration = MinImageAnimationDuration;
            var ppEntryEffect = Document.SlideTransitionPool.RandomEntryEffect();
            Slide.SlideShowTransition.EntryEffect = ppEntryEffect;
            Slide.SlideShowTransition.Duration = 1;
            if (!string.IsNullOrEmpty(imagePath))
            {
                PrimaryImage = Slide.Shapes.AddPicture(imagePath,
                    MsoTriState.msoTrue, MsoTriState.msoTrue, 0, 0);
                //宽 / 高
                var sizeRatio = PrimaryImage.Width/PrimaryImage.Height;
                var screenRatio = Document.SlideWidth/Document.SlideHeight;
                // 图片比屏幕更高一些，纵向。
                isPrimaryImageVertical = sizeRatio < screenRatio;
                // 先尝试放大图片，使其一边超出屏幕。
                if (isPrimaryImageVertical)
                    PrimaryImage.Width = Document.SlideWidth;
                else
                    PrimaryImage.Height = Document.SlideHeight;
                // 计算超出的部分。
                var scrollOffsetRelative = isPrimaryImageVertical
                    ? PrimaryImage.Height/Document.SlideHeight
                    : PrimaryImage.Width/Document.SlideWidth; // > 1.0
                if (scrollOffsetRelative < 1.2)
                {
                    switch (rnd.Next(0, 3))
                    {
                        case 0:
                            ImageAnimation(PrimaryImageAnimation.Fit);
                            break;
                        case 1:
                            ImageAnimation(PrimaryImageAnimation.Expand);
                            break;
                        default:
                            ImageAnimation(PrimaryImageAnimation.Shrink);
                            break;
                    }
                }
                else if (isPrimaryImageVertical)
                    ImageAnimation(PrimaryImageAnimation.ScrollNear);
                else
                    ImageAnimation(PrimaryImageAnimation.ScrollFar);
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

        public AnimationInfo AddAnimation(Shape shape, MsoAnimEffect effect,
            TimeSpan delay, TimeSpan duration, AnimationOptions options)
        {
            var e = Slide.TimeLine.MainSequence.AddEffect(shape, effect,
                trigger: MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            e.Exit = (options & AnimationOptions.Exit) == AnimationOptions.Exit
                ? MsoTriState.msoTrue
                : MsoTriState.msoFalse;
            var ani =
                new AnimationInfo((((options & AnimationOptions.WithPrevious) == AnimationOptions.WithPrevious
                    ? Animations.LastOrDefault()?.StartAt
                    : Animations.LastOrDefault()?.EndAt) ?? TimeSpan.Zero) + delay,
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
            textBox.TextFrame.TextRange.Text = content;
            textBox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            textBox.TextFrame.TextRange.Font.Size = SubtitleFontSize;
            textBox.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            textBox.TextFrame.TextRange.Font.Shadow = MsoTriState.msoTrue;
            textBox.TextFrame2.TextRange.Font.Shadow.Transparency = 0;
            textBox.TextFrame2.TextRange.Font.Line.Visible = MsoTriState.msoTrue;
            textBox.TextFrame2.TextRange.Font.Line.ForeColor.RGB = 0;
            textBox.TextFrame2.TextRange.Font.Line.Weight = 1;
            return textBox;
        }

        protected override void OnLeavingClosure()
        {
            base.OnLeavingClosure();
            Slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
            var animationTime = Animations.Count > 0 ? Animations.Last().EndAt : TimeSpan.Zero;
            Slide.SlideShowTransition.AdvanceTime = Math.Max(PrimaryImageAnimationDuration, (float)animationTime.TotalSeconds) +
                                                    PagePersistTime;
        }

        private string GenerateScrollMotionPath(bool scrollFar)
        {
            if (PrimaryImage == null) return null;
            if (isPrimaryImageVertical)
            {
                // 纵向
                PrimaryImage.Width = Document.SlideWidth;
                var scrollOffsetRelative = (PrimaryImage.Height - Document.SlideHeight) / Document.SlideHeight;
                if (scrollFar) scrollOffsetRelative = -scrollOffsetRelative;
                PrimaryImage.Top = scrollFar
                    ? 0
                    : Document.SlideHeight - PrimaryImage.Height;
                return $"M 0 0 L 0 {scrollOffsetRelative}";
            }
            else
            {
                // 横向
                PrimaryImage.Height = Document.SlideHeight;
                var scrollOffsetRelative = (PrimaryImage.Width - Document.SlideWidth) / Document.SlideWidth;
                if (scrollFar) scrollOffsetRelative = -scrollOffsetRelative;
                PrimaryImage.Left = scrollFar
                    ? 0
                    : Document.SlideWidth - PrimaryImage.Width;
                return $"M 0 0 L {scrollOffsetRelative} 0";
            }
        }

        /// <summary>
        /// 以图形的中心为参考点，放大/缩小图形。
        /// </summary>
        public static void ScaleFromCenter(Shape shape, float ratio)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            shape.Left -= shape.Width*(ratio - 1)/2;
            shape.Top -= shape.Width*(ratio - 1)/2;
            shape.Width *= ratio;
            shape.Height *= ratio;
        }

        public Closure Persist(float persistTime)
        {
            PagePersistTime = persistTime;
            return this;
        }

        private void AlignImage(bool allowCrop, bool alignCenter)
        {
            // allowCrop == true 放大图片，使其一边超出屏幕。
            // allowCrop == false 放大图片，但不要使其超出屏幕。
            if (isPrimaryImageVertical ^ !allowCrop)
                PrimaryImage.Width = Document.SlideWidth;
            else
                PrimaryImage.Height = Document.SlideHeight;
            if (alignCenter)
            {
                PrimaryImage.Left = (Document.SlideWidth - PrimaryImage.Width)/2;
                PrimaryImage.Top = (Document.SlideHeight - PrimaryImage.Height)/2;
            }
            else
            {
                PrimaryImage.Left = 0;
                PrimaryImage.Top = 0;
            }
        }

        private void SubstitutePrimaryImageAnimation(MsoAnimEffect? effect)
        {
            if (effect == null)
            {
                if (primaryImageAnimation != null)
                {
                    primaryImageAnimation.Delete();
                    primaryImageAnimation = null;
                }
                return;
            }
            if (primaryImageAnimation == null)
            {
                primaryImageAnimation = Slide.TimeLine.MainSequence.AddEffect(PrimaryImage,
                    effectId: effect.Value,
                    trigger: MsoAnimTriggerType.msoAnimTriggerWithPrevious, Index: 1);
            }
            else
            {
                primaryImageAnimation.EffectType = effect.Value;
            }
            primaryImageAnimation.Timing.Duration = PrimaryImageAnimationDuration;
            primaryImageAnimation.Timing.SmoothEnd = MsoTriState.msoTrue;
        }

        private readonly Random rnd = new Random();

        [ClosureOperation]
        public Closure ImageAnimation(PrimaryImageAnimation animation)
        {
            if (PrimaryImage == null) return this;
            if (animation == PrimaryImageAnimation.ExpandOrShink)
                animation = rnd.Next(0, 2) == 0 ? PrimaryImageAnimation.Expand : PrimaryImageAnimation.Shrink;
            switch (animation)
            {
                case PrimaryImageAnimation.None:
                    AlignImage(false, true);
                    SubstitutePrimaryImageAnimation(null);
                    break;
                case PrimaryImageAnimation.Fit:
                    AlignImage(true, true);
                    SubstitutePrimaryImageAnimation(null);
                    break;
                case PrimaryImageAnimation.Expand:
                    AlignImage(true, true);
                    PrimaryImageAnimationDuration = MinImageAnimationDuration;
                    SubstitutePrimaryImageAnimation(MsoAnimEffect.msoAnimEffectGrowShrink);
                    var b = primaryImageAnimation.Behaviors[1];
                    b.ScaleEffect.ByX = b.ScaleEffect.ByY = 105; //注意单位是%
                    break;
                case PrimaryImageAnimation.Shrink:
                    AlignImage(true, true);
                    ScaleFromCenter(PrimaryImage, 1.05f);
                    PrimaryImageAnimationDuration = MinImageAnimationDuration;
                    SubstitutePrimaryImageAnimation(MsoAnimEffect.msoAnimEffectGrowShrink);
                    b = primaryImageAnimation.Behaviors[1];
                    b.ScaleEffect.ByX = b.ScaleEffect.ByY = 100/1.05f; //注意单位是%
                    break;
                case PrimaryImageAnimation.ScrollNear:
                case PrimaryImageAnimation.ScrollFar:
                    AlignImage(true, false);
                    // 滚动过长的图片。
                    // 计算超出的部分。
                    var scrollOffsetRelative = isPrimaryImageVertical
                        ? PrimaryImage.Height / Document.SlideHeight
                        : PrimaryImage.Width / Document.SlideWidth; // > 1.0
                    if (Math.Abs(scrollOffsetRelative) > 2)
                    {
                        // 如果超出屏幕的区域太长……
                        PrimaryImageAnimationDuration = MinImageAnimationDuration*Math.Abs(scrollOffsetRelative);
                    }
                    else
                    {
                        PrimaryImageAnimationDuration = MinImageAnimationDuration;
                    }
                    SubstitutePrimaryImageAnimation(MsoAnimEffect.msoAnimEffectCustom);
                    b = primaryImageAnimation.Behaviors.Add(MsoAnimType.msoAnimTypeMotion);
                    b.MotionEffect.Path = GenerateScrollMotionPath(animation == PrimaryImageAnimation.ScrollFar);
                    primaryImageAnimation.Timing.Duration = PrimaryImageAnimationDuration;
                    primaryImageAnimation.Timing.SmoothEnd = MsoTriState.msoCTrue;
                    break;
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
        /// 插入BGM，注意，此功能仅用作打样。建议后期手动合成音轨。
        /// </summary>
        [ClosureOperation]
        public Closure Music(string path = "", int stopAfterSlides = 999)
        {
            var media = Slide.Shapes.AddMediaObject2(Path.Combine(Document.WorkPath, path));
            var effectIndex = Slide.TimeLine.MainSequence.AddEffect(media, MsoAnimEffect.msoAnimEffectMediaPlay,
                trigger: MsoAnimTriggerType.msoAnimTriggerWithPrevious).Index;
            media.Left = -media.Width;
            media.Top = -media.Height;
            Func<Effect> effect = () => Slide.TimeLine.MainSequence[effectIndex];
            ////注意下面两行的顺序不能反。
            //effect().EffectInformation.PlaySettings.HideWhileNotPlaying = MsoTriState.msoTrue;
            effect().EffectInformation.PlaySettings.StopAfterSlides = stopAfterSlides;
            return this;
        }

        [ClosureOperation]
        public Closure Transition(PpEntryEffect effect = PpEntryEffect.ppEffectNone)
        {
            Slide.SlideShowTransition.EntryEffect = effect;
            return this;
        }

        public PageClosure(Closure parent, Slide slide) : base(parent)
        {
            Slide = slide;
        }
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
        WithPrevious = 8,
    };

    /// <summary>
    /// 一张图片的滚动方向（视图滚动方向）。
    /// </summary>
    public enum ImageScrollDirection
    {
        LeftUp,
        RightDown
    }
}