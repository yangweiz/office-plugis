using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Suite.Core.Commands;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Suite.PowerPoint.Services
{
    public sealed class ShapeService
    {
        public CommandResult AlignLeft(PowerPoint.Selection selection)
        {
            return Execute(selection, 2, "左对齐", delegate(PowerPoint.ShapeRange shapes)
            {
                float left;
                PowerPoint.Shape first = shapes[1];
                try { left = first.Left; }
                finally { Release(first); }
                for (int index = 2; index <= shapes.Count; index++) SetPosition(shapes[index], left, null);
                return CommandResult.Success("已将 " + shapes.Count + " 个形状按首个形状左对齐。");
            });
        }

        public CommandResult AlignRight(PowerPoint.Selection selection)
        {
            return Execute(selection, 2, "右对齐", delegate(PowerPoint.ShapeRange shapes)
            {
                float right;
                PowerPoint.Shape first = shapes[1];
                try { right = first.Left + first.Width; }
                finally { Release(first); }
                for (int index = 2; index <= shapes.Count; index++)
                {
                    PowerPoint.Shape shape = shapes[index];
                    try { shape.Left = right - shape.Width; }
                    finally { Release(shape); }
                }
                return CommandResult.Success("已将 " + shapes.Count + " 个形状按首个形状右对齐。");
            });
        }

        public CommandResult AlignTop(PowerPoint.Selection selection)
        {
            return Execute(selection, 2, "上对齐", delegate(PowerPoint.ShapeRange shapes)
            {
                float top;
                PowerPoint.Shape first = shapes[1];
                try { top = first.Top; }
                finally { Release(first); }
                for (int index = 2; index <= shapes.Count; index++) SetPosition(shapes[index], null, top);
                return CommandResult.Success("已将 " + shapes.Count + " 个形状按首个形状上对齐。");
            });
        }

        public CommandResult AlignBottom(PowerPoint.Selection selection)
        {
            return Execute(selection, 2, "下对齐", delegate(PowerPoint.ShapeRange shapes)
            {
                float bottom;
                PowerPoint.Shape first = shapes[1];
                try { bottom = first.Top + first.Height; }
                finally { Release(first); }
                for (int index = 2; index <= shapes.Count; index++)
                {
                    PowerPoint.Shape shape = shapes[index];
                    try { shape.Top = bottom - shape.Height; }
                    finally { Release(shape); }
                }
                return CommandResult.Success("已将 " + shapes.Count + " 个形状按首个形状下对齐。");
            });
        }

        public CommandResult FlipHorizontal(PowerPoint.Selection selection)
        {
            return Flip(selection, MsoFlipCmd.msoFlipHorizontal, "水平翻转");
        }

        public CommandResult FlipVertical(PowerPoint.Selection selection)
        {
            return Flip(selection, MsoFlipCmd.msoFlipVertical, "垂直翻转");
        }

        public CommandResult UniformWidth(PowerPoint.Selection selection)
        {
            return Execute(selection, 2, "统一宽度", delegate(PowerPoint.ShapeRange shapes)
            {
                float width;
                PowerPoint.Shape first = shapes[1];
                try { width = first.Width; }
                finally { Release(first); }
                for (int index = 2; index <= shapes.Count; index++) SetSize(shapes[index], width, null);
                return CommandResult.Success("已将 " + shapes.Count + " 个形状的宽度统一为首个形状的宽度。");
            });
        }

        public CommandResult UniformHeight(PowerPoint.Selection selection)
        {
            return Execute(selection, 2, "统一高度", delegate(PowerPoint.ShapeRange shapes)
            {
                float height;
                PowerPoint.Shape first = shapes[1];
                try { height = first.Height; }
                finally { Release(first); }
                for (int index = 2; index <= shapes.Count; index++) SetSize(shapes[index], null, height);
                return CommandResult.Success("已将 " + shapes.Count + " 个形状的高度统一为首个形状的高度。");
            });
        }

        public CommandResult ApplyThemeColor(PowerPoint.Selection selection, MsoThemeColorIndex themeColor, string colorName)
        {
            return Execute(selection, 1, "设置主题填充", delegate(PowerPoint.ShapeRange shapes)
            {
                for (int index = 1; index <= shapes.Count; index++)
                {
                    PowerPoint.Shape shape = null;
                    PowerPoint.FillFormat fill = null;
                    ColorFormat foreground = null;
                    try
                    {
                        shape = shapes[index];
                        fill = shape.Fill;
                        fill.Visible = MsoTriState.msoTrue;
                        fill.Solid();
                        foreground = fill.ForeColor;
                        foreground.ObjectThemeColor = themeColor;
                        foreground.Brightness = 0;
                    }
                    finally
                    {
                        Release(foreground);
                        Release(fill);
                        Release(shape);
                    }
                }
                return CommandResult.Success("已仅将 " + shapes.Count + " 个选定形状的填充设置为" + colorName + "。");
            });
        }

        private static CommandResult Flip(PowerPoint.Selection selection, MsoFlipCmd direction, string operation)
        {
            return Execute(selection, 1, operation, delegate(PowerPoint.ShapeRange shapes)
            {
                for (int index = 1; index <= shapes.Count; index++)
                {
                    PowerPoint.Shape shape = null;
                    try
                    {
                        shape = shapes[index];
                        shape.Flip(direction);
                    }
                    finally { Release(shape); }
                }
                return CommandResult.Success("已" + operation + " " + shapes.Count + " 个选定形状。");
            });
        }

        private static void SetPosition(PowerPoint.Shape shape, float? left, float? top)
        {
            try
            {
                if (left.HasValue) shape.Left = left.Value;
                if (top.HasValue) shape.Top = top.Value;
            }
            finally { Release(shape); }
        }

        private static void SetSize(PowerPoint.Shape shape, float? width, float? height)
        {
            try
            {
                if (width.HasValue) shape.Width = width.Value;
                if (height.HasValue) shape.Height = height.Value;
            }
            finally { Release(shape); }
        }

        private static CommandResult Execute(PowerPoint.Selection selection, int minimumCount, string operation,
            Func<PowerPoint.ShapeRange, CommandResult> action)
        {
            PowerPoint.ShapeRange shapes = null;
            CommandResult blocked = GetSelectedShapes(selection, minimumCount, operation, out shapes);
            if (blocked != null) return blocked;
            try { return action(shapes); }
            finally { Release(shapes); }
        }

        private static CommandResult GetSelectedShapes(PowerPoint.Selection selection, int minimumCount, string operation, out PowerPoint.ShapeRange shapes)
        {
            shapes = null;
            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                return CommandResult.Blocked(operation + "需要在幻灯片编辑视图中选择形状；当前没有有效形状选区。");

            shapes = selection.ShapeRange;
            if (shapes == null || shapes.Count < minimumCount)
            {
                Release(shapes);
                shapes = null;
                return CommandResult.Blocked(minimumCount == 1
                    ? operation + "至少需要选择一个形状，操作未执行。"
                    : operation + "至少需要选择两个形状，并以首个形状为基准；操作未执行。");
            }
            return null;
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
    }
}
