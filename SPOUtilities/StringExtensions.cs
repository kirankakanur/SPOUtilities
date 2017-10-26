using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOUtilities
{
    using System.Threading;

    /// <summary>
    /// String extension methods
    /// </summary>

    public static class StringExtensions
    {
        /// <summary>
        /// Strips the img source.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        public static string StripImgSrc(this string value)
        {
            if (String.IsNullOrEmpty(value)) return String.Empty;

            if (value.ToLowerInvariant().StartsWith("<img"))
            {
                var start = value.IndexOf("src=") + 5;
                if (start == -1) return String.Empty;

                var len = value.IndexOf("\"", start) - start;

                if (len == -1) len = value.IndexOf("'", start) - start;

                if (len > -1) return value.Substring(start, len);
            }

            return value;
        }

        /// <summary>
        /// To title case.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        public static string ToTitleCase(this string value)
        {
            return Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(value.ToLower());
        }

        /// <summary>
        /// Makes the query string friendly.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        public static string MakeQueryStringFriendly(this string value)
        {
            return System.Web.HttpUtility.UrlEncode(value);
        }

        /// <summary>
        /// Truncates to last character.
        /// </summary>
        /// <param name="valueToTruncate">The value to truncate.</param>
        /// <param name="maxLength">The maximum length.</param>
        /// <param name="addElipsis">if set to <c>true</c> [add elipsis].</param>
        /// <returns></returns>
        public static string TruncateToLastCharacter(this string valueToTruncate, int maxLength, bool addElipsis)
        {
            //make room for ellipsis
            if (addElipsis && maxLength > 3)
                maxLength = maxLength - 3;

            if (valueToTruncate == null)
            {
                return String.Empty;
            }

            if (valueToTruncate.Length <= maxLength)
            {
                return valueToTruncate;
            }
            string retValue = valueToTruncate.Remove(maxLength);

            if (addElipsis)
            {
                retValue += "&hellip;";
            }
            return retValue;
        }

        /// <summary>
        /// Truncates to last space.
        /// </summary>
        /// <param name="valueToTruncate">The value to truncate.</param>
        /// <param name="maxLength">The maximum length.</param>
        /// <param name="addElipsis">if set to <c>true</c> [add elipsis].</param>
        /// <returns></returns>
        public static string TruncateToLastSpace(this string valueToTruncate, int maxLength, bool addElipsis)
        {
            if (valueToTruncate == null)
            {
                return String.Empty;
            }

            if (valueToTruncate.Length <= maxLength)
            {
                return valueToTruncate;
            }

            string retValue = valueToTruncate;

            int lastSpaceIndex = retValue.LastIndexOf(" ",
                                                      maxLength, StringComparison.CurrentCultureIgnoreCase);

            if (lastSpaceIndex > -1)
            {
                retValue = retValue.Remove(lastSpaceIndex);
            }

            if (retValue.Length < valueToTruncate.Length && addElipsis)
            {
                retValue += "&hellip;";
            }
            return retValue;
        }

        /// <summary>
        /// Truncates to last space.
        /// </summary>
        /// <param name="valueToTruncate">The value to truncate.</param>
        /// <param name="maxLength">The maximum length.</param>
        /// <returns></returns>
        public static string TruncateToLastSpace(this string valueToTruncate, int maxLength)
        {
            return TruncateToLastSpace(valueToTruncate, maxLength, true);
        }

        /// <summary>
        /// Removes all HTML tags using HTML AgilityPack
        /// </summary>
        /// <param name="html">The HTML.</param>
        /// <returns></returns>
        //public static string RemoveAllHtmlTags(this string html)
        //{
        //    var doc = new HtmlDocument();
        //    doc.LoadHtml(html);

        //    return doc.DocumentNode.InnerText;
        //}

        /// <summary>
        /// This method removes ALL style and class attributes from all tags in a string of Html.
        /// </summary>
        /// <param name="html"></param>
        /// <returns></returns>
        //public static string RemoveStyleAndClassAttributesFromHtml(this string html)
        //{
        //    if (html.StartsWith("<img"))
        //    {
        //        if (!html.EndsWith("</img>") && !html.EndsWith("/>"))
        //            html = html.TrimEnd('>');
        //        html += "/>";
        //    }
        //    html = html.Replace("&", "&amp;");
        //    always add root node as html blocks don't always have root.
        //    var preCleaned = html.Insert(0, "<rootnode>");
        //    preCleaned += "</rootnode>";

        //    var xmlDoc = XDocument.Parse(preCleaned);

        //    xmlDoc.Descendants().Attributes("style").Remove();
        //    xmlDoc.Descendants().Attributes("class").Remove();

        //    var cleanXML = xmlDoc.ToString();

        //    var output = cleanXML.Replace("<rootnode>", string.Empty).Replace("</rootnode>", string.Empty);

        //    return output.Replace("&amp;", "&");
        //}

        /// <summary>
        /// Gets the left half.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <param name="splitter">The splitter.</param>
        /// <returns></returns>
        public static string GetLeftHalf(this string str, char splitter)
        {
            string[] parts = str.Split(splitter);
            return parts[0];

        }

        /// <summary>
        /// Gets the left half.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <param name="splitter">The splitter.</param>
        /// <returns></returns>
        public static string GetLeftHalf(this string str, string splitter)
        {
            var parts = str.Split(new string[] { splitter }, StringSplitOptions.RemoveEmptyEntries);

            return parts[0];
        }

        /// <summary>
        /// Gets the right half.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <param name="splitter">The splitter.</param>
        /// <returns></returns>
        public static string GetRightHalf(this string str, char splitter)
        {
            string[] parts = str.Split(splitter);
            int count = parts.Count();
            if (count > 1)
                return parts[count - 1];
            return str;
        }

        /// <summary>
        /// Gets the right half.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <param name="splitter">The splitter.</param>
        /// <returns></returns>
        public static string GetRightHalf(this string str, string splitter)
        {
            string[] parts = str.Split(new[] { splitter }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Count() > 1)
                return parts[parts.Count() - 1];
            return str;
        }

        /// <summary>
        /// Splits the and remove empty entries.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <param name="splitter">The splitter.</param>
        /// <returns></returns>
        public static string[] SplitAndRemoveEmptyEntries(this string str, string splitter)
        {
            return str.Split(splitter.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
        }

        /// <summary>
        /// To safe string.
        /// </summary>
        /// <param name="obj">The object.</param>
        /// <returns></returns>
        public static string ToSafeString(this object obj)
        {
            return (obj ?? String.Empty).ToString();
        }

        /// <summary>
        /// Deserializes a JSON object into list of strings, removing any blanks
        /// </summary>
        /// <param name="audiences">A serialized list of strings in JSON format</param>
        /// <returns>List of strings</returns>
        //public static IList<string> DeserializeStringListNoBlanks(this string audiences)
        //{
        //    var deserializedFilters = new JavaScriptSerializer().Deserialize<List<string>>(audiences);
        //    deserializedFilters = deserializedFilters.Where(i => !string.IsNullOrEmpty(i)).ToList();
        //    return deserializedFilters;
        //}
    }
}
