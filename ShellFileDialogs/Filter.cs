using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ShellFileDialogs.Native;

namespace ShellFileDialogs
{
    public class Filter
    {
        private static readonly char[] _semiColon = { ';' };
        private static readonly char[] _pipe = { '|' };

        /// <summary></summary>
        /// <param name="displayName">Required. Cannot be <see langword="null" />, empty, nor whitespace.</param>
        /// <param name="extensions">
        ///     Required. Cannot be <see langword="null" /> or empty (i.e. at least one extension filter must
        ///     be specified).
        /// </param>
        public Filter(string displayName, params string[] extensions)
            : this(displayName, (IEnumerable<string>)extensions)
        {
        }

        /// <summary></summary>
        /// <param name="displayName">Required. Cannot be <see langword="null" />, empty, nor whitespace.</param>
        /// <param name="extensions">
        ///     Required. The collection cannot be <see langword="null" /> or empty - and it cannot contain
        ///     null, empty, or whitespace values - nor duplicate values.
        /// </param>
        /// <exception cref="ArgumentNullException">
        ///     Thrown when <paramref name="displayName" /> or <paramref name="extensions" />
        ///     is <see langword="null" />.
        /// </exception>
        /// <exception cref="ArgumentException">
        ///     Thrown when <paramref name="extensions" /> is empty or contains only empty file
        ///     extensions.
        /// </exception>
        public Filter(string displayName, IEnumerable<string> extensions)
        {
            if (string.IsNullOrWhiteSpace(displayName)) throw new ArgumentNullException(nameof(displayName));
            if (extensions is null) throw new ArgumentNullException(nameof(extensions));

            DisplayName = displayName.Trim();

            Extensions = extensions
                .Select(s => s.Trim()) // Trim whitespace
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .ToList(); // make a copy to prevent possible changes

            if (Extensions.Count == 0)
                throw new ArgumentException(
                    "Extensions collection must not be empty, nor can it contain only null, empty or whitespace extensions.",
                    nameof(extensions));
        }

        public string DisplayName { get; }

        /// <summary>
        ///     All extension values have their leading dot and any leading asterisk filter trimmed, so if &quot;<c>*.wav</c>
        ///     &quot; is passed as an extension string to <see cref="Filter.Filter(string, string[])" /> then it will appear in
        ///     this list as &quot;<c>wav</c>&quot;.
        /// </summary>
        public IReadOnlyList<string> Extensions { get; }

        /// <summary>Returns <see langword="null" /> if the string couldn't be parsed.</summary>
        public static IReadOnlyList<Filter>? ParseWindowsFormsFilter(string filter)
        {
            // https://msdn.microsoft.com/en-us/library/system.windows.forms.filedialog.filter(v=vs.110).aspx
            if (string.IsNullOrWhiteSpace(filter)) return null;

            var components = filter.Split(_pipe, StringSplitOptions.RemoveEmptyEntries);
            if (components.Length % 2 != 0) return null;

            var filters = new Filter[components.Length / 2];
            var fi = 0;
            for (var i = 0; i < components.Length; i += 2)
            {
                var displayName = components[i];
                var extensionsCat = components[i + 1];

                var extensions = extensionsCat.Split(_semiColon, StringSplitOptions.RemoveEmptyEntries);

                filters[fi] = new Filter(displayName, extensions);
                fi++;
            }

            return filters;
        }

        private string ToFilterSpecString()
        {
            var sb = new StringBuilder();
            var first = true;
            foreach (var extension in Extensions)
            {
                if (!first) _ = sb.Append(';');
                first = false;

                _ = sb.Append("*.");
                _ = sb.Append(extension);
            }

            return sb.ToString();
        }

        private void ToExtensionList(StringBuilder sb)
        {
            var first = true;
            foreach (var extension in Extensions)
            {
                if (!first) _ = sb.Append(", ");
                first = false;

                _ = sb.Append("*.");
                _ = sb.Append(extension);
            }
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            _ = sb.Append(DisplayName);

            _ = sb.Append(" (");
            ToExtensionList(sb);
            _ = sb.Append(')');

            return sb.ToString();
        }

        internal FilterSpec ToFilterSpec()
        {
            var filter = ToFilterSpecString();
            return new FilterSpec(DisplayName, filter);
        }
    }
}