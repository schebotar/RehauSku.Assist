﻿using System.Collections.Generic;
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
            Regex regex = new Regex(@"\d{2}.\d{2}.\d{2}");

            if (!regex.IsMatch(line))
                return line;

            string match = regex.Match(line).Value;

            int endFaceA = int.Parse($"{match[0]}{match[1]}"),                
                side = int.Parse($"{match[3]}{match[4]}"),
                endFaceB = int.Parse($"{match[6]}{match[7]}");

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
            string modifiedMatch = string.Join(" ", additions) + piece;

            return line.Replace(match, modifiedMatch);
        }
    }
}