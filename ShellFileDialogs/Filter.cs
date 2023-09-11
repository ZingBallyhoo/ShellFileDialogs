using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ShellFileDialogs
{
    public class Filter
    {
        public string DisplayName { get; }

        /// <summary>
        ///     All extension values have their leading dot and any leading asterisk filter trimmed, so if &quot;<c>*.wav</c>
        ///     &quot; is passed as an extension string to <see cref="Filter.Filter(string, string[])" /> then it will appear in
        ///     this list as &quot;<c>wav</c>&quot;.
        /// </summary>
        public IReadOnlyList<string> Extensions { get; }
        
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

        public string ToFilterSpecString()
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
    }
}