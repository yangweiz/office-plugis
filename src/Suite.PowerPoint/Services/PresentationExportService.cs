using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Suite.Core.Batching;
using Suite.Core.Commands;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Suite.PowerPoint.Services
{
    public sealed class PresentationExportService
    {
        public CommandResult ExportSlidesAsPng(PowerPoint.Presentation presentation, string outputDirectory)
        {
            CommandResult readiness = ValidatePresentation(presentation);
            if (readiness != null) return readiness;
            if (string.IsNullOrWhiteSpace(outputDirectory)) return CommandResult.Blocked("请选择有效的图片导出目录。");

            string fullOutputDirectory;
            try { fullOutputDirectory = Path.GetFullPath(outputDirectory); }
            catch (Exception ex) { return CommandResult.Failed("导出目录路径无效：" + ex.Message); }
            if (!Directory.Exists(fullOutputDirectory)) return CommandResult.Failed("导出目录不存在或当前不可访问：" + fullOutputDirectory);

            PowerPoint.Slides slides = null;
            PowerPoint.PageSetup pageSetup = null;
            try
            {
                slides = presentation.Slides;
                if (slides.Count == 0) return CommandResult.Blocked("当前演示文稿没有可导出的幻灯片。");
                string presentationName = MakeSafeFileName(Path.GetFileNameWithoutExtension(presentation.Name));
                pageSetup = presentation.PageSetup;
                int width = Math.Max(1, (int)Math.Round(pageSetup.SlideWidth * 96f / 72f));
                int height = Math.Max(1, (int)Math.Round(pageSetup.SlideHeight * 96f / 72f));
                int exported = 0;

                for (int index = 1; index <= slides.Count; index++)
                {
                    PowerPoint.Slide slide = null;
                    int slideIndex = index;
                    try
                    {
                        slide = slides[index];
                        slideIndex = slide.SlideIndex;
                        string fileName = presentationName + "_幻灯片_" + slideIndex.ToString("000") + ".png";
                        string outputPath = OutputPathPolicy.GetAvailablePath(Path.Combine(fullOutputDirectory, fileName));
                        slide.Export(outputPath, "PNG", width, height);
                        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
                            throw new IOException("PowerPoint 未生成有效的 PNG 文件。");
                        exported++;
                    }
                    catch (Exception ex)
                    {
                        string error = "导出第 " + slideIndex + " 张幻灯片时失败：" + ex.Message;
                        if (exported == 0) return CommandResult.Failed(error);
                        return CommandResult.Partial(error + " 此前已成功导出 " + exported + " 张，位置：" + fullOutputDirectory);
                    }
                    finally { Release(slide); }
                }

                return CommandResult.Success("已逐页导出 " + exported + " 张 PNG 图片；同名文件通过追加序号保留，未覆盖已有文件。", fullOutputDirectory);
            }
            finally
            {
                Release(pageSetup);
                Release(slides);
            }
        }

        private static CommandResult ValidatePresentation(PowerPoint.Presentation presentation)
        {
            if (presentation == null) return CommandResult.Blocked("请先打开要导出的演示文稿。");
            if (string.IsNullOrWhiteSpace(presentation.Path)) return CommandResult.Blocked("演示文稿尚未保存到本地路径，请先保存再导出。");
            if (presentation.Saved != MsoTriState.msoTrue) return CommandResult.Blocked("演示文稿包含未保存的更改，请先保存再导出。");
            if (!Directory.Exists(presentation.Path) || !File.Exists(presentation.FullName))
                return CommandResult.Blocked("演示文稿的保存路径无效或文件不存在，请重新保存到有效本地路径后再导出。");
            return null;
        }

        private static string MakeSafeFileName(string name)
        {
            char[] invalid = Path.GetInvalidFileNameChars();
            string safe = new string((name ?? string.Empty).Select(character => invalid.Contains(character) ? '_' : character).ToArray()).Trim().TrimEnd('.', ' ');
            return safe.Length == 0 ? "演示文稿" : safe;
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
    }
}
