
using System;
using System.Linq;
using System.Text.RegularExpressions;

//https://qiita.com/omochi_motimoti/items/b6261e3cacda11d460dd

public static class StringExtensions
{
    /// <summary>
    /// 文字列をExcelのカラム番号へ変換します
    /// </summary>
    /// <param name="column">a-z/A-Zのみで構成された文字列</param>
    public static int ToColumnNumber(this string source)
    {
        if (string.IsNullOrWhiteSpace(source))
            return 0;

        string buf = source.ToUpper();
        if (Regex.IsMatch(buf, @"^[A-Z]+$") == false)
            throw new FormatException("Argument format is only alphabet");

        // 変換後がint.MaxValueに収まるか?
        //if (buf.CompareTo("FXSHRXW") >= 1)
        if (source.Length >= "FXSHRXW".Length && buf.CompareTo("FXSHRXW") >= 1)
            throw new ArgumentOutOfRangeException("Argument range max \"FXSHRXW\"");

        return buf.ToColumnNumber(0);
    }

    /// <summary>
    /// 文字列→Excelカラム番号変換処理の実体
    /// </summary>
    /// <param name="column">A-Zのみで構成された文字列</param>
    /// <param name="call_num">呼び出し回数</param>
    static int ToColumnNumber(this string source, int call_num)
    {
        if (source == "") return 0;
        int digit = (int)Math.Pow(26, call_num);
        return ((source.Last() - 'A' + 1) * digit) + ToColumnNumber(source.Substring(0, source.Length - 1), ++call_num);
    }
}

public static class IntExtensions
{
    /// <summary>
    /// 数値をExcelのカラム文字へ変換します
    /// </summary>
    /// <param name="column"></param>
    /// <returns></returns>
    public static string ToColumnName(this int source)
    {
        if (source < 1) return string.Empty;
        return ToColumnName((source - 1) / 26) + (char)('A' + ((source - 1) % 26));
    }
}