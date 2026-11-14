using System;
using System.Runtime.InteropServices;
using Suite.Core.Commands;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Suite.PowerPoint.Services
{
    public sealed class PresentationCleanupService
    {
        public CommandResult RemoveBlankSlides(PowerPoint.Presentation presentation)
        {
            if (presentation == null) return CommandResult.Blocked("请先打开要清理的演示文稿。");
            int removed = 0;
            PowerPoint.Slides slides = null;
            try
            {
                slides = presentation.Slides;
                for (int index = slides.Count; index >= 1; index--)
                {
                    PowerPoint.Slide slide = null;
                    try
                    {
                        slide = slides[index];
                        if (!IsBlankSlide(slide)) continue;
                        slide.Delete();
                        removed++;
                    }
                    finally { Release(slide); }
                }
            }
            finally { Release(slides); }

            return CommandResult.Success(removed == 0
                ? "未发现可安全清理的空白页；隐藏页及其他含内容页面均未更改。"
                : "已清理 " + removed + " 张空白页；隐藏页、媒体、备注、隐藏形状及其他内容均已保留。");
        }

        public CommandResult RemoveHiddenSlides(PowerPoint.Presentation presentation)
        {
            if (presentation == null) return CommandResult.Blocked("请先打开要清理的演示文稿。");
            int removed = 0;
            PowerPoint.Slides slides = null;
            try
            {
                slides = presentation.Slides;
                for (int index = slides.Count; index >= 1; index--)
                {
                    PowerPoint.Slide slide = null;
                    try
                    {
                        slide = slides[index];
                        if (!IsHidden(slide)) continue;
                        slide.Delete();
                        removed++;
                    }
                    finally { Release(slide); }
                }
            }
            finally { Release(slides); }
            return CommandResult.Success(removed == 0 ? "当前演示文稿没有隐藏页。" : "已按明确操作删除 " + removed + " 张隐藏页。");
        }

        public CommandResult RemoveAnimations(PowerPoint.Presentation presentation)
        {
            if (presentation == null) return CommandResult.Blocked("请先打开要清理的演示文稿。");
            int removed = 0;
            PowerPoint.Slides slides = null;
            try
            {
                slides = presentation.Slides;
                for (int slideIndex = 1; slideIndex <= slides.Count; slideIndex++)
                {
                    PowerPoint.Slide slide = null;
                    PowerPoint.TimeLine timeLine = null;
                    PowerPoint.Sequence mainSequence = null;
                    PowerPoint.Sequences interactiveSequences = null;
                    try
                    {
                        slide = slides[slideIndex];
                        timeLine = slide.TimeLine;
                        mainSequence = timeLine.MainSequence;
                        removed += DeleteEffects(mainSequence);
                        interactiveSequences = timeLine.InteractiveSequences;
                        for (int sequenceIndex = interactiveSequences.Count; sequenceIndex >= 1; sequenceIndex--)
                        {
                            PowerPoint.Sequence sequence = null;
                            try
                            {
                                sequence = interactiveSequences[sequenceIndex];
                                removed += DeleteEffects(sequence);
                            }
                            finally { Release(sequence); }
                        }
                    }
                    finally
                    {
                        Release(interactiveSequences);
                        Release(mainSequence);
                        Release(timeLine);
                        Release(slide);
                    }
                }
            }
            finally { Release(slides); }
            return CommandResult.Success(removed == 0 ? "未发现需要清理的动画。" : "已清理 " + removed + " 个动画效果。");
        }

        public CommandResult RemoveTransitions(PowerPoint.Presentation presentation)
        {
            if (presentation == null) return CommandResult.Blocked("请先打开要清理的演示文稿。");
            int changed = 0;
            PowerPoint.Slides slides = null;
            try
            {
                slides = presentation.Slides;
                for (int index = 1; index <= slides.Count; index++)
                {
                    PowerPoint.Slide slide = null;
                    PowerPoint.SlideShowTransition transition = null;
                    try
                    {
                        slide = slides[index];
                        transition = slide.SlideShowTransition;
                        bool needsCleanup = transition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectNone
                            || transition.AdvanceOnTime == MsoTriState.msoTrue
                            || transition.AdvanceTime > 0;
                        if (!needsCleanup) continue;
                        transition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
                        transition.AdvanceOnTime = MsoTriState.msoFalse;
                        transition.AdvanceTime = 0;
                        changed++;
                    }
                    finally
                    {
                        Release(transition);
                        Release(slide);
                    }
                }
            }
            finally { Release(slides); }
            return CommandResult.Success(changed == 0
                ? "未发现需要清理的切换效果或自动换片时间。"
                : "已清理 " + changed + " 张幻灯片的切换效果和自动换片时间；隐藏状态及单击换片设置未更改。");
        }

        public bool IsBlankSlide(PowerPoint.Slide slide)
        {
            if (slide == null) return false;
            PowerPoint.Shapes shapes = null;
            PowerPoint.Comments comments = null;
            PowerPoint.Tags tags = null;
            PowerPoint.Hyperlinks hyperlinks = null;
            PowerPoint.TimeLine timeLine = null;
            PowerPoint.Sequence mainSequence = null;
            PowerPoint.Sequences interactiveSequences = null;
            PowerPoint.SlideShowTransition transition = null;
            try
            {
                if (IsHidden(slide)) return false;
                shapes = slide.Shapes;
                if (shapes.Count > 0 && HasSlideShapeContent(shapes)) return false;
                comments = slide.Comments;
                tags = slide.Tags;
                hyperlinks = slide.Hyperlinks;
                if (comments.Count > 0 || tags.Count > 0 || hyperlinks.Count > 0) return false;
                if (slide.FollowMasterBackground == MsoTriState.msoFalse) return false;
                timeLine = slide.TimeLine;
                mainSequence = timeLine.MainSequence;
                interactiveSequences = timeLine.InteractiveSequences;
                if (mainSequence.Count > 0 || interactiveSequences.Count > 0) return false;
                transition = slide.SlideShowTransition;
                if (transition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectNone || transition.AdvanceOnTime == MsoTriState.msoTrue) return false;
                return !HasNotesContent(slide);
            }
            finally
            {
                Release(transition);
                Release(interactiveSequences);
                Release(mainSequence);
                Release(timeLine);
                Release(hyperlinks);
                Release(tags);
                Release(comments);
                Release(shapes);
            }
        }

        private static int DeleteEffects(PowerPoint.Sequence sequence)
        {
            int removed = 0;
            for (int index = sequence.Count; index >= 1; index--)
            {
                PowerPoint.Effect effect = null;
                try
                {
                    effect = sequence[index];
                    effect.Delete();
                    removed++;
                }
                finally { Release(effect); }
            }
            return removed;
        }

        private static bool IsHidden(PowerPoint.Slide slide)
        {
            PowerPoint.SlideShowTransition transition = null;
            try
            {
                transition = slide.SlideShowTransition;
                return transition.Hidden == MsoTriState.msoTrue;
            }
            finally { Release(transition); }
        }

        private static bool HasSlideShapeContent(PowerPoint.Shapes shapes)
        {
            for (int index = 1; index <= shapes.Count; index++)
            {
                PowerPoint.Shape shape = null;
                try
                {
                    shape = shapes[index];
                    if (shape.Visible == MsoTriState.msoFalse) return true;
                    if (shape.Type == MsoShapeType.msoMedia) return true;
                    if (shape.Type != MsoShapeType.msoPlaceholder) return true;
                    if (HasVisibleText(shape)) return true;
                    if (!IsEmptyTextPlaceholder(shape)) return true;
                }
                finally { Release(shape); }
            }
            return false;
        }

        private static bool IsEmptyTextPlaceholder(PowerPoint.Shape shape)
        {
            PowerPoint.PpPlaceholderType type = GetPlaceholderType(shape);
            return type == PowerPoint.PpPlaceholderType.ppPlaceholderTitle
                || type == PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle
                || type == PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle
                || type == PowerPoint.PpPlaceholderType.ppPlaceholderBody
                || type == PowerPoint.PpPlaceholderType.ppPlaceholderVerticalTitle
                || type == PowerPoint.PpPlaceholderType.ppPlaceholderVerticalBody;
        }

        private static bool HasNotesContent(PowerPoint.Slide slide)
        {
            if (slide.HasNotesPage != MsoTriState.msoTrue) return false;
            PowerPoint.Slide notesPage = null;
            PowerPoint.Shapes shapes = null;
            try
            {
                notesPage = slide.NotesPage;
                shapes = notesPage.Shapes;
                for (int index = 1; index <= shapes.Count; index++)
                {
                    PowerPoint.Shape shape = null;
                    try
                    {
                        shape = shapes[index];
                        if (HasVisibleText(shape)) return true;
                        if (shape.Type != MsoShapeType.msoPlaceholder) return true;
                        PowerPoint.PpPlaceholderType type = GetPlaceholderType(shape);
                        bool canBeEmpty = type == PowerPoint.PpPlaceholderType.ppPlaceholderBody
                            || type == PowerPoint.PpPlaceholderType.ppPlaceholderHeader
                            || type == PowerPoint.PpPlaceholderType.ppPlaceholderFooter
                            || type == PowerPoint.PpPlaceholderType.ppPlaceholderDate
                            || type == PowerPoint.PpPlaceholderType.ppPlaceholderSlideNumber
                            || type == PowerPoint.PpPlaceholderType.ppPlaceholderObject;
                        if (!canBeEmpty) return true;
                    }
                    finally { Release(shape); }
                }
                return false;
            }
            finally
            {
                Release(shapes);
                Release(notesPage);
            }
        }

        private static PowerPoint.PpPlaceholderType GetPlaceholderType(PowerPoint.Shape shape)
        {
            PowerPoint.PlaceholderFormat placeholder = null;
            try
            {
                placeholder = shape.PlaceholderFormat;
                return placeholder.Type;
            }
            finally { Release(placeholder); }
        }

        private static bool HasVisibleText(PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame != MsoTriState.msoTrue) return false;
            PowerPoint.TextFrame textFrame = null;
            PowerPoint.TextRange textRange = null;
            try
            {
                textFrame = shape.TextFrame;
                if (textFrame.HasText != MsoTriState.msoTrue) return false;
                textRange = textFrame.TextRange;
                return !string.IsNullOrWhiteSpace(textRange.Text);
            }
            finally
            {
                Release(textRange);
                Release(textFrame);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
    }
}
