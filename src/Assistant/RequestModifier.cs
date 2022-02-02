using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace RehauSku.Assistant
{
    static class RequestModifier
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

            return replace._tPieceNormalize();
        }

        private static string _tPieceNormalize(this string line)
        {
            Regex regex = new Regex(@"\d{2}.\d{2}.\d{2}");

            if (!regex.IsMatch(line))
                return line;

            string match = regex.Match(line).Value;

            int side = int.Parse($"{match[3]}{match[4]}");
            int[] endFaces = new int[]
            {
                int.Parse($"{match[0]}{match[1]}"),
                int.Parse($"{match[6]}{match[7]}")
            };

            if (new[] { endFaces[0], endFaces[1], side }.Any(x => x == 45 || x == 90 || x == 87))
                return line;

            List<string> additions = new List<string>();

            if (endFaces.All(x => x < side))
                additions.Add("увеличенный боковой");

            else
            {
                if (new[] { endFaces[0], endFaces[1], side }.Distinct().Count() == 1)
                    additions.Add("равнопроходной");
                else
                    additions.Add("уменьшенный");

                if (endFaces.Any(x => x > side))
                    additions.Add("боковой");

                if (endFaces[0] != endFaces[1])
                    additions.Add("торцевой");
            }

            string piece = $" {endFaces.Max()}-{side}-{endFaces.Min()} ";
            string modifiedMatch = string.Join(" ", additions) + piece;

            return line.Replace(match, modifiedMatch);
        }
    }
}