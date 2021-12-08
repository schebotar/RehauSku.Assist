using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Rehau.Sku.Assist
{
    public static class RequestModifier
    {
        public static string CleanRequest(this string input)
        {
            string replace = new StringBuilder(input)
                .Replace("+", " plus ")
                .Replace("РХ", "")
                .Replace("º", " ")
                .Replace(".", " ")
                .Replace("Ø", " ")
                .ToString();

            return replace._tPiece();
        }

        private static string _tPiece(this string line)
        {
            if (!line.ToLower().Contains("тройник"))
                return line;

            string m = Regex.Match(line, @"\d{2}.\d{2}.\d{2}").Value;

            int endFaceA = int.Parse($"{m[0]}{m[1]}");
            int side = int.Parse($"{m[3]}{m[4]}");
            int endFaceB = int.Parse($"{m[6]}{m[7]}");

            int[] endFaces = new[] { endFaceA, endFaceB };

            List<string> additions = new List<string>();

            if (endFaces.All(x => x < side))
                additions.Add("увеличенный боковой");

            else
            {
                if (new[] { endFaceA, endFaceB, side }.Distinct().Count() == 1)
                    additions.Add("равнопроходной");
                else
                    additions.Add("уменьшенный");

                if (endFaces.Any(x => x > side))
                    additions.Add("боковой");
                if (endFaceA != endFaceB)
                    additions.Add("торцевой");
            }

            string piece = $" {endFaces.Max()}-{side}-{endFaces.Min()} ";
            string replace = string.Join(" ", additions) + piece;

            return line.Replace(m, replace);
        }
    }
}