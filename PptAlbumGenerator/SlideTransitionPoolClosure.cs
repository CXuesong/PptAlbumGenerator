﻿using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;

namespace PptAlbumGenerator
{
    /// <summary>
    /// 用于管理可用的幻灯片切换效果。
    /// </summary>
    internal class SlideTransitionPoolClosure : Closure
    {
        /// <summary>
        /// 获取所有可用的幻灯片切换效果。（不包括 ppEffectRandom 。）
        /// </summary>
        private static PpEntryEffect[] AvailableSlideEntryEffects =
        {
            PpEntryEffect.ppEffectCut,
            PpEntryEffect.ppEffectCutThroughBlack,
            //PpEntryEffect.ppEffectRandom,
            PpEntryEffect.ppEffectBlindsHorizontal,
            PpEntryEffect.ppEffectBlindsVertical,
            PpEntryEffect.ppEffectCheckerboardAcross,
            PpEntryEffect.ppEffectCheckerboardDown,
            PpEntryEffect.ppEffectCoverLeft,
            PpEntryEffect.ppEffectCoverUp,
            PpEntryEffect.ppEffectCoverRight,
            PpEntryEffect.ppEffectCoverDown,
            PpEntryEffect.ppEffectCoverLeftUp,
            PpEntryEffect.ppEffectCoverRightUp,
            PpEntryEffect.ppEffectCoverLeftDown,
            PpEntryEffect.ppEffectCoverRightDown,
            PpEntryEffect.ppEffectDissolve,
            PpEntryEffect.ppEffectFade,
            PpEntryEffect.ppEffectUncoverLeft,
            PpEntryEffect.ppEffectUncoverUp,
            PpEntryEffect.ppEffectUncoverRight,
            PpEntryEffect.ppEffectUncoverDown,
            PpEntryEffect.ppEffectUncoverLeftUp,
            PpEntryEffect.ppEffectUncoverRightUp,
            PpEntryEffect.ppEffectUncoverLeftDown,
            PpEntryEffect.ppEffectUncoverRightDown,
            PpEntryEffect.ppEffectRandomBarsHorizontal,
            PpEntryEffect.ppEffectRandomBarsVertical,
            PpEntryEffect.ppEffectStripsUpLeft,
            PpEntryEffect.ppEffectStripsUpRight,
            PpEntryEffect.ppEffectStripsDownLeft,
            PpEntryEffect.ppEffectStripsDownRight,
            PpEntryEffect.ppEffectStripsLeftUp,
            PpEntryEffect.ppEffectStripsRightUp,
            PpEntryEffect.ppEffectStripsLeftDown,
            PpEntryEffect.ppEffectStripsRightDown,
            PpEntryEffect.ppEffectWipeLeft,
            PpEntryEffect.ppEffectWipeUp,
            PpEntryEffect.ppEffectWipeRight,
            PpEntryEffect.ppEffectWipeDown,
            PpEntryEffect.ppEffectBoxOut,
            PpEntryEffect.ppEffectBoxIn,
            ////[UNS]PpEntryEffect.ppEffectFlyFromLeft,
            ////[UNS]PpEntryEffect.ppEffectFlyFromTop,
            ////[UNS]PpEntryEffect.ppEffectFlyFromRight,
            ////[UNS]PpEntryEffect.ppEffectFlyFromBottom,
            ////[UNS]PpEntryEffect.ppEffectFlyFromTopLeft,
            ////[UNS]PpEntryEffect.ppEffectFlyFromTopRight,
            ////[UNS]PpEntryEffect.ppEffectFlyFromBottomLeft,
            ////[UNS]PpEntryEffect.ppEffectFlyFromBottomRight,
            ////[UNS]PpEntryEffect.ppEffectPeekFromLeft,
            ////[UNS]PpEntryEffect.ppEffectPeekFromDown,
            ////[UNS]PpEntryEffect.ppEffectPeekFromRight,
            ////[UNS]PpEntryEffect.ppEffectPeekFromUp,
            ////[UNS]PpEntryEffect.ppEffectCrawlFromLeft,
            ////[UNS]PpEntryEffect.ppEffectCrawlFromUp,
            ////[UNS]PpEntryEffect.ppEffectCrawlFromRight,
            ////[UNS]PpEntryEffect.ppEffectCrawlFromDown,
            ////[UNS]PpEntryEffect.ppEffectZoomIn,
            ////[UNS]PpEntryEffect.ppEffectZoomInSlightly,
            ////[UNS]PpEntryEffect.ppEffectZoomOut,
            ////[UNS]PpEntryEffect.ppEffectZoomOutSlightly,
            ////[UNS]PpEntryEffect.ppEffectZoomCenter,
            ////[UNS]PpEntryEffect.ppEffectZoomBottom,
            ////[UNS]PpEntryEffect.ppEffectStretchAcross,
            ////[UNS]PpEntryEffect.ppEffectStretchLeft,
            ////[UNS]PpEntryEffect.ppEffectStretchUp,
            ////[UNS]PpEntryEffect.ppEffectStretchRight,
            ////[UNS]PpEntryEffect.ppEffectStretchDown,
            ////[UNS]PpEntryEffect.ppEffectSwivel,
            ////[UNS]PpEntryEffect.ppEffectSpiral,
            PpEntryEffect.ppEffectSplitHorizontalOut,
            PpEntryEffect.ppEffectSplitHorizontalIn,
            PpEntryEffect.ppEffectSplitVerticalOut,
            PpEntryEffect.ppEffectSplitVerticalIn,
            ////[UNS]PpEntryEffect.ppEffectFlashOnceFast,
            ////[UNS]PpEntryEffect.ppEffectFlashOnceMedium,
            ////[UNS]PpEntryEffect.ppEffectFlashOnceSlow,
            //PpEntryEffect.ppEffectAppear,
            PpEntryEffect.ppEffectCircleOut,
            PpEntryEffect.ppEffectDiamondOut,
            PpEntryEffect.ppEffectCombHorizontal,
            PpEntryEffect.ppEffectCombVertical,
            PpEntryEffect.ppEffectFadeSmoothly,
            PpEntryEffect.ppEffectNewsflash,
            PpEntryEffect.ppEffectPlusOut,
            PpEntryEffect.ppEffectPushDown,
            PpEntryEffect.ppEffectPushLeft,
            PpEntryEffect.ppEffectPushRight,
            PpEntryEffect.ppEffectPushUp,
            PpEntryEffect.ppEffectWedge,
            PpEntryEffect.ppEffectWheel1Spoke,
            PpEntryEffect.ppEffectWheel2Spokes,
            PpEntryEffect.ppEffectWheel3Spokes,
            PpEntryEffect.ppEffectWheel4Spokes,
            PpEntryEffect.ppEffectWheel8Spokes,
            PpEntryEffect.ppEffectWheelReverse1Spoke,
            PpEntryEffect.ppEffectVortexLeft,
            PpEntryEffect.ppEffectVortexUp,
            PpEntryEffect.ppEffectVortexRight,
            PpEntryEffect.ppEffectVortexDown,
            PpEntryEffect.ppEffectRippleCenter,
            PpEntryEffect.ppEffectRippleRightUp,
            PpEntryEffect.ppEffectRippleLeftUp,
            PpEntryEffect.ppEffectRippleLeftDown,
            PpEntryEffect.ppEffectRippleRightDown,
            PpEntryEffect.ppEffectGlitterDiamondLeft,
            PpEntryEffect.ppEffectGlitterDiamondUp,
            PpEntryEffect.ppEffectGlitterDiamondRight,
            PpEntryEffect.ppEffectGlitterDiamondDown,
            PpEntryEffect.ppEffectGlitterHexagonLeft,
            PpEntryEffect.ppEffectGlitterHexagonUp,
            PpEntryEffect.ppEffectGlitterHexagonRight,
            PpEntryEffect.ppEffectGlitterHexagonDown,
            PpEntryEffect.ppEffectGalleryLeft,
            PpEntryEffect.ppEffectGalleryRight,
            PpEntryEffect.ppEffectConveyorLeft,
            PpEntryEffect.ppEffectConveyorRight,
            PpEntryEffect.ppEffectDoorsVertical,
            PpEntryEffect.ppEffectDoorsHorizontal,
            PpEntryEffect.ppEffectWindowVertical,
            PpEntryEffect.ppEffectWindowHorizontal,
            PpEntryEffect.ppEffectWarpIn,
            PpEntryEffect.ppEffectWarpOut,
            PpEntryEffect.ppEffectFlyThroughIn,
            PpEntryEffect.ppEffectFlyThroughOut,
            PpEntryEffect.ppEffectFlyThroughInBounce,
            PpEntryEffect.ppEffectFlyThroughOutBounce,
            PpEntryEffect.ppEffectRevealSmoothLeft,
            PpEntryEffect.ppEffectRevealSmoothRight,
            PpEntryEffect.ppEffectRevealBlackLeft,
            PpEntryEffect.ppEffectRevealBlackRight,
            PpEntryEffect.ppEffectHoneycomb,
            PpEntryEffect.ppEffectFerrisWheelLeft,
            PpEntryEffect.ppEffectFerrisWheelRight,
            PpEntryEffect.ppEffectSwitchLeft,
            PpEntryEffect.ppEffectSwitchUp,
            PpEntryEffect.ppEffectSwitchRight,
            PpEntryEffect.ppEffectSwitchDown,
            PpEntryEffect.ppEffectFlipLeft,
            PpEntryEffect.ppEffectFlipUp,
            PpEntryEffect.ppEffectFlipRight,
            PpEntryEffect.ppEffectFlipDown,
            PpEntryEffect.ppEffectFlashbulb,
            PpEntryEffect.ppEffectShredStripsIn,
            PpEntryEffect.ppEffectShredStripsOut,
            PpEntryEffect.ppEffectShredRectangleIn,
            PpEntryEffect.ppEffectShredRectangleOut,
            PpEntryEffect.ppEffectCubeLeft,
            PpEntryEffect.ppEffectCubeUp,
            PpEntryEffect.ppEffectCubeRight,
            PpEntryEffect.ppEffectCubeDown,
            PpEntryEffect.ppEffectRotateLeft,
            PpEntryEffect.ppEffectRotateUp,
            PpEntryEffect.ppEffectRotateRight,
            PpEntryEffect.ppEffectRotateDown,
            PpEntryEffect.ppEffectBoxLeft,
            PpEntryEffect.ppEffectBoxUp,
            PpEntryEffect.ppEffectBoxRight,
            PpEntryEffect.ppEffectBoxDown,
            PpEntryEffect.ppEffectOrbitLeft,
            PpEntryEffect.ppEffectOrbitUp,
            PpEntryEffect.ppEffectOrbitRight,
            PpEntryEffect.ppEffectOrbitDown,
            PpEntryEffect.ppEffectPanLeft,
            PpEntryEffect.ppEffectPanUp,
            PpEntryEffect.ppEffectPanRight,
            PpEntryEffect.ppEffectPanDown,
        };

        private IList<PpEntryEffect> _Transitions = AvailableSlideEntryEffects;

        public SlideTransitionPoolClosure(Closure parent) : base(parent)
        {
        }

        public IList<PpEntryEffect> Transitions => _Transitions;

        private void EnsureLocalCopy()
        {
            if (_Transitions == AvailableSlideEntryEffects)
                _Transitions = new List<PpEntryEffect>(AvailableSlideEntryEffects);
        }

        [ClosureOperation]
        public Closure Add(PpEntryEffect effect)
        {
            EnsureLocalCopy();
            _Transitions.Add(effect);
            return this;
        }

        [ClosureOperation]
        public Closure Remove(PpEntryEffect effect)
        {
            EnsureLocalCopy();
            _Transitions.Remove(effect);
            return this;
        }

        [ClosureOperation]
        public Closure Clear(PpEntryEffect effect)
        {
            if (_Transitions == AvailableSlideEntryEffects)
                _Transitions = new List<PpEntryEffect>();
            else
                _Transitions.Clear();
            return this;
        }

        private Random rnd = new Random();

        public PpEntryEffect RandomEntryEffect()
        {
            if (_Transitions.Count == 0) return PpEntryEffect.ppEffectNone;
            return _Transitions[rnd.Next(0, _Transitions.Count)];
        }
    }
}