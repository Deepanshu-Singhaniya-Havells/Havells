using System;
using System.Text.RegularExpressions;

namespace Havells.Dataverse.CustomConnector
{
    public class APValidate
    {
        public static bool IsValidMobileNumber(string mobileNumber)
        {
            // Regex pattern to match numbers starting with 6, 7, 8, or 9 and having exactly 10 digits
            string pattern = @"^[6-9]\d{9}$";
            if (string.IsNullOrWhiteSpace(mobileNumber))
            {
                return false;
            }
            if (mobileNumber.Length > 10)
            {
                mobileNumber = mobileNumber.Substring(mobileNumber.Length - 10, 10);
            }
            // Check if the mobile number is not null or empty and matches the pattern
            if (Regex.IsMatch(mobileNumber, pattern))
            {
                return true;
            }
            return false;
        }
        public static bool IsValidEmail(string emailid)
        {
            try
            {
                var regx = new Regex(@"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$");
                if (string.IsNullOrWhiteSpace(emailid))
                    return false;
                if (regx.IsMatch(emailid))
                {
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }
        public static bool IsValidString(string input)
        {
            // Check if the input is null, empty, or whitespace
            if (string.IsNullOrWhiteSpace(input))
            {
                return false;
            }
            try
            {
                // Regular expression to match only letters (uppercase and lowercase with spaces)
                var regexItem = new Regex(@"^[a-zA-Z\s]+$");
                // Check if the input matches the regex
                return regexItem.IsMatch(input);
            }

            catch (Exception)

            {
                return false;
            }

        }
        public static bool IsValidboolen(string input)
        {
            if (string.IsNullOrEmpty(input))
                return false;
            try
            {
                bool.Parse(input);
            }
            catch (FormatException)
            {
                return false;
            }
            return true;
        }
        public static bool isAlphaNumeric(string serialno)
        {
            try
            {
                Regex rg = new Regex(@"^[a-zA-Z0-9]*$");
                return rg.IsMatch(serialno);
            }
            catch { return false; }
        }
        public static bool IsvalidGuid(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return false;
            }
            try
            {
                Guid.Parse(input);
            }
            catch (FormatException)
            {
                return false;
            }
            return true;
        }
        public static bool IsDecimal(string input)
        {
            decimal A = Convert.ToDecimal(0);
            if (string.IsNullOrWhiteSpace(input) || decimal.Compare(Convert.ToDecimal(input), A) == 0)
            {
                return false;
            }
            try
            {
                decimal.Parse(input);
            }
            catch (FormatException)
            {
                return false;
            }
            return true;
        }
        public static bool IsInteger(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return false;
            }
            try
            {
                Int32.Parse(input);
            }
            catch (FormatException)
            {
                return false;
            }
            return true;
        }
        public static bool IsvalidDate(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return false;
            }
            try
            {
                DateTime.Parse(input);
            }
            catch (FormatException)
            {
                return false;
            }
            return true;
        }
        public static bool NumericValue(string Value)
        {
            var regx = new Regex("^\\d+$");
            if (string.IsNullOrWhiteSpace(Value))
            {
                return false;
            }
            try
            {
                return regx.Match(Value).Success;
            }
            catch (FormatException)
            {
                return false;
            }
        }
        public static bool IsNumeric(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return false;
            }
            try
            {
                double.Parse(input);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }
        public static string ValidateTwoDates(DateTime ToDate, DateTime FromDate)
        {
            if (ToDate >= FromDate)
            {
                return "";
            }
            return "To Date should be greater than or equal to From Date.";
        }
        public static bool IsNumericGreaterThanZero(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return false;
            }
            try
            {
                double result = double.Parse(input);
                if (result == 0)
                {
                    return false;
                }
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }
    }
    public static class OptionSetValues
    {
        internal static readonly int[] SalutationItem = new int[] { 1, 2, 3, 4, 5 };
        internal static readonly int[] GenderCodeItem = new int[] { 1, 2 };
    }
}
