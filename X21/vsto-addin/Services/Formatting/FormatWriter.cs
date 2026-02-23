using System;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using X21.Logging;
using X21.Models;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Service responsible for applying formatting to Excel ranges
    /// </summary>
    public class FormatWriter
    {
        /// <summary>
        /// Applies formatting settings to a range
        /// </summary>
        public void ApplyFormattingToRange(Range targetRange, FormatSettings format)
        {
            try
            {
                Logger.Info($"Applying formatting to range: {targetRange.Address[false, false]}");

                // Apply number format
                if (!string.IsNullOrEmpty(format?.NumberFormat))
                {
                    Logger.Info($"Applying number format: {format.NumberFormat}");
                    targetRange.NumberFormat = format.NumberFormat;
                }

                // Apply font formatting
                if (format?.Bold.HasValue == true)
                {
                    Logger.Info($"Applying bold: {format.Bold.Value}");
                    targetRange.Font.Bold = format.Bold.Value;
                }

                if (format?.Italic.HasValue == true)
                {
                    Logger.Info($"Applying italic: {format.Italic.Value}");
                    targetRange.Font.Italic = format.Italic.Value;
                }

                if (format?.Underline.HasValue == true)
                {
                    Logger.Info($"Applying underline: {format.Underline.Value}");
                    targetRange.Font.Underline = format.Underline.Value ? XlUnderlineStyle.xlUnderlineStyleSingle : XlUnderlineStyle.xlUnderlineStyleNone;
                }

                if (format?.FontSize.HasValue == true)
                {
                    Logger.Info($"Applying font size: {format.FontSize.Value}");
                    targetRange.Font.Size = format.FontSize.Value;
                }

                if (!string.IsNullOrEmpty(format?.FontName))
                {
                    Logger.Info($"Applying font name: {format.FontName}");
                    targetRange.Font.Name = format.FontName;
                }

                if (!string.IsNullOrEmpty(format?.FontColor))
                {
                    Logger.Info($"Applying font color: {format.FontColor}");
                    targetRange.Font.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(format.FontColor));
                }

                // Apply background color
                if (!string.IsNullOrEmpty(format?.BackgroundColor))
                {
                    Logger.Info($"Applying background color: {format.BackgroundColor}");
                    if (format.BackgroundColor == "none")
                    {
                        targetRange.Interior.Color = XlColorIndex.xlColorIndexNone; // Clear background
                    }
                    else
                    {
                        targetRange.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(format.BackgroundColor));
                    }
                }

                // Apply alignment
                if (!string.IsNullOrEmpty(format?.Alignment))
                {
                    Logger.Info($"Applying alignment: {format.Alignment}");
                    switch (format.Alignment.ToLower())
                    {
                        case "center":
                            targetRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            break;
                        case "right":
                            targetRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                            break;
                        case "left":
                        case "default":
                            targetRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                            break;
                    }
                }

                if (format?.ClearBorders == true)
                {
                    Logger.Info("Clearing borders");
                    targetRange.Borders.LineStyle = XlLineStyle.xlLineStyleNone;
                }

                // Apply borders - COMMENTED OUT TO STRIP BORDERS
                // if (format?.Border != null)
                // {
                //     Logger.Info("Applying detailed border settings");
                //     ApplyBorderSettingsToRange(targetRange, format.Border);
                // }

                Logger.Info("Formatting applied successfully");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                throw new Exception($"Error applying formatting to range {targetRange.Address[false, false]}: {ex.Message}");
            }
        }

        /*         /// <summary>
                /// Applies border settings to a range
                /// </summary>
                public void ApplyBorderSettingsToRange(Range targetRange, BorderSettings borderSettings)
                {
                    if (borderSettings == null) return;

                    try
                    {
                        var borders = targetRange.Borders;

                        // Apply top border
                        if (!string.IsNullOrEmpty(borderSettings.Top) && borderSettings.Top != "none")
                        {
                            borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                            borders[XlBordersIndex.xlEdgeTop].Weight = GetBorderWeightFromString(borderSettings.Top);
                        }
                        else if (borderSettings.Top == "none")
                        {
                            borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlLineStyleNone;
                        }

                        // Apply bottom border
                        if (!string.IsNullOrEmpty(borderSettings.Bottom) && borderSettings.Bottom != "none")
                        {
                            borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                            borders[XlBordersIndex.xlEdgeBottom].Weight = GetBorderWeightFromString(borderSettings.Bottom);
                        }
                        else if (borderSettings.Bottom == "none")
                        {
                            borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlLineStyleNone;
                        }

                        // Apply left border
                        if (!string.IsNullOrEmpty(borderSettings.Left) && borderSettings.Left != "none")
                        {
                            borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                            borders[XlBordersIndex.xlEdgeLeft].Weight = GetBorderWeightFromString(borderSettings.Left);
                        }
                        else if (borderSettings.Left == "none")
                        {
                            borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlLineStyleNone;
                        }

                        // Apply right border
                        if (!string.IsNullOrEmpty(borderSettings.Right) && borderSettings.Right != "none")
                        {
                            borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                            borders[XlBordersIndex.xlEdgeRight].Weight = GetBorderWeightFromString(borderSettings.Right);
                        }
                        else if (borderSettings.Right == "none")
                        {
                            borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlLineStyleNone;
                        }

                        Logger.Info($"Applied border settings to range {targetRange.Address}");
                    }
                    catch (Exception ex)
                    {
                        Logger.Info($"Error applying border settings to range {targetRange.Address}: {ex.Message}");
                        throw;
                    }
                }
         */
        /*         /// <summary>
                /// Converts border weight string to Excel border weight enum
                /// </summary>
                private XlBorderWeight GetBorderWeightFromString(string weight)
                {
                    return weight?.ToLower() switch
                    {
                        "thin" => XlBorderWeight.xlThin,
                        "medium" => XlBorderWeight.xlMedium,
                        "thick" => XlBorderWeight.xlThick,
                        _ => XlBorderWeight.xlThin
                    };
                }
            } */
    }
}
