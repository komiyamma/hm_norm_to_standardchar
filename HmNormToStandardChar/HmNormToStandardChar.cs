using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace HmNormToStandardChar;

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
[Guid("45C935C4-A35E-49DE-8373-D058AE8574C8")]
public class HmNormToStandardChar
{
    public string Normalize(string text)
    {
        try
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            var lines = text.Split('\n');

            Parallel.For(0, lines.Length, i =>
            {
                lines[i] = NormalizeLine(lines[i]);
            });

            return String.Join("\n", lines);
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }

    private Encoding encoder = Encoding.GetEncoding("Shift_JIS");
    private bool CanConvertSJIS(string text)
    {
        try
        {
            var encodedBytes = encoder.GetBytes(text);
            var decodedText = encoder.GetString(encodedBytes);
            return text == decodedText;
        }
        catch (Exception e)
        {
            return false;
        }
    }

    private string NormalizeLine(string text)
    {
        if (CanConvertSJIS(text))
        {
            return text;
        }

        var walker = System.Globalization.StringInfo.GetTextElementEnumerator(text);

        var dstText = "";
        while(walker.MoveNext()) {
            var chr = walker.GetTextElement();
            dstText += NormalizeCharacter(chr);
        }
        return dstText;
    }

    private string NormalizeCharacter(string text)
    {
        if (CanConvertSJIS(text))
        {
            return text;
        }

        var textNFC = text.Normalize(System.Text.NormalizationForm.FormC);
        if (CanConvertSJIS(textNFC))
        {
            return textNFC;
        }
        if (text != textNFC)
        {
            return textNFC;
        }

        return text.Normalize(System.Text.NormalizationForm.FormKC);
    }
}
