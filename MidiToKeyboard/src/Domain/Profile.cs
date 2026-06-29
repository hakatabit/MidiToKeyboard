using System;
using System.Collections.Generic;

namespace MidiToKeyboard.Domain
{
    public sealed class Profile
    {
        private readonly Dictionary<int, char> noteMappings;

        public string Name { get; }

        public IReadOnlyDictionary<int, char> NoteMappings
        {
            get { return noteMappings; }
        }

        public Profile(string name, IDictionary<int, char> noteMappings)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("Profile name must not be null, empty, or whitespace.", nameof(name));
            }

            if (noteMappings == null)
            {
                throw new ArgumentNullException(nameof(noteMappings));
            }

            Name = name;
            this.noteMappings = new Dictionary<int, char>(noteMappings);
        }
    }
}
