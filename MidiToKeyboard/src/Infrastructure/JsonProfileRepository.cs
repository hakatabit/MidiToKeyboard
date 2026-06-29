using MidiToKeyboard.Application;
using MidiToKeyboard.Domain;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace MidiToKeyboard.Infrastructure
{
    public sealed class JsonProfileRepository : IProfileRepository
    {
        private readonly string filePath;

        public JsonProfileRepository(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("File path must not be null, empty, or whitespace.", nameof(filePath));
            }

            this.filePath = filePath;
        }

        public Profile Load(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("Profile name must not be null, empty, or whitespace.", nameof(name));
            }

            MappingSet set = LoadMappingSets()
                .Where(mappingSet => mappingSet != null)
                .FirstOrDefault(mappingSet => string.Equals(mappingSet.name, name, StringComparison.Ordinal));

            if (set == null)
            {
                throw new KeyNotFoundException("Profile was not found: " + name);
            }

            return CreateProfile(set);
        }

        public IEnumerable<string> ListNames()
        {
            return LoadMappingSets()
                .Where(mappingSet => mappingSet != null)
                .Select(mappingSet => mappingSet.name);
        }

        private List<MappingSet> LoadMappingSets()
        {
            using (var mappingsStream = File.OpenRead(filePath))
            {
                MappingFile mappingFile = JsonSerializer.Deserialize<MappingFile>(mappingsStream);
                return mappingFile?.sets ?? new List<MappingSet>();
            }
        }

        private static Profile CreateProfile(MappingSet set)
        {
            var noteMappings = new Dictionary<int, char>();

            if (set.mappings != null)
            {
                foreach (var mapping in set.mappings)
                {
                    int noteNumber;
                    if (!int.TryParse(mapping.Key, out noteNumber))
                    {
                        continue;
                    }

                    if (string.IsNullOrEmpty(mapping.Value))
                    {
                        continue;
                    }

                    noteMappings[noteNumber] = mapping.Value[0];
                }
            }

            return new Profile(set.name, noteMappings);
        }

        private class MappingFile
        {
            public List<MappingSet> sets { get; set; }
        }

        private class MappingSet
        {
            public string name { get; set; }
            public Dictionary<string, string> mappings { get; set; }
        }
    }
}
