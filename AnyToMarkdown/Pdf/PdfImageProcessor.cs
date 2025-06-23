using SkiaSharp;

namespace AnyToMarkdown.Pdf;

internal static class PdfImageProcessor
{
    public static string ConvertImageToBase64(ReadOnlySpan<byte> imageBytes, int quality = 75)
    {
        try
        {
            using var bitmap = SKBitmap.Decode(imageBytes);
            using var encodedData = bitmap.Encode(SKEncodedImageFormat.Png, quality);
            if (encodedData == null)
            {
                return string.Empty;
            }

            byte[] pngBytes = encodedData.ToArray();
            string base64 = Convert.ToBase64String(pngBytes);
            return $"data:image/png;base64,{base64}";
        }
        catch
        {
            return string.Empty;
        }
    }
}