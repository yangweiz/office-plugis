using System;
using System.Linq;

namespace Suite.Core.Data
{
    public static class ChineseIdCard
    {
        private static readonly int[] Weights = { 7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2 };
        private const string CheckCodes = "10X98765432";

        public static bool IsValid(string value)
        {
            DateTime ignored;
            return TryGetBirthDate(value, out ignored) && HasValidCheckCode(value);
        }

        public static bool TryGetBirthDate(string value, out DateTime birthDate)
        {
            birthDate = default(DateTime);
            if (string.IsNullOrWhiteSpace(value)) return false;
            value = value.Trim();
            if (value.Length != 18 || !value.Take(17).All(char.IsDigit) || !(char.IsDigit(value[17]) || value[17] == 'X' || value[17] == 'x')) return false;
            if (value.Substring(0, 6) == "000000" || value.Substring(12, 3) == "000") return false;
            int year, month, day;
            if (!int.TryParse(value.Substring(6, 4), out year) || !int.TryParse(value.Substring(10, 2), out month) || !int.TryParse(value.Substring(12, 2), out day)) return false;
            try { birthDate = new DateTime(year, month, day); }
            catch (ArgumentOutOfRangeException) { return false; }
            return birthDate <= DateTime.Today;
        }

        public static string GetGender(string value)
        {
            if (!IsValid(value)) throw new ArgumentException("无效身份证号码。", "value");
            return (value[16] - '0') % 2 == 0 ? "女" : "男";
        }

        public static int GetAge(string value, DateTime asOfDate)
        {
            DateTime birth;
            if (!TryGetBirthDate(value, out birth) || !HasValidCheckCode(value)) throw new ArgumentException("无效身份证号码。", "value");
            if (asOfDate.Date < birth) throw new ArgumentException("年龄计算日期不得早于出生日期。", "asOfDate");
            int age = asOfDate.Year - birth.Year;
            if (asOfDate.Date < birth.AddYears(age)) age--;
            return age;
        }

        public static string Mask(string value)
        {
            if (!IsValid(value)) throw new ArgumentException("无效身份证号码。", "value");
            return value.Substring(0, 6) + new string('*', 8) + value.Substring(14, 4);
        }

        private static bool HasValidCheckCode(string value)
        {
            int sum = 0;
            for (int index = 0; index < 17; index++) sum += (value[index] - '0') * Weights[index];
            return char.ToUpperInvariant(value[17]) == CheckCodes[sum % 11];
        }
    }
}
