using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using MidiToKeyboard.Application;
using MidiToKeyboard.Domain;

namespace MidiToKeyboard.Infrastructure
{
    /// <summary>
    /// JSON ファイルからのプロファイル読み込み
    /// </summary>
    public sealed class JsonProfileRepository : IProfileRepository
    {
        private readonly string _filePath;

        /// <summary>
        /// 読み込み対象の JSON ファイルを指定して初期化
        /// </summary>
        public JsonProfileRepository(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException(
                    "ファイルパスに null、空文字列、または空白のみの文字列は指定できません。",
                    nameof(filePath));
            }

            _filePath = filePath;
        }

        /// <inheritdoc />
        public Profile Load(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException(
                    "プロファイル名に null、空文字列、または空白のみの文字列は指定できません。",
                    nameof(name));
            }

            MappingSet set = LoadMappingSets()
                .Where(mappingSet => mappingSet != null)
                .FirstOrDefault(mappingSet => string.Equals(
                    mappingSet.name,
                    name,
                    StringComparison.Ordinal));

            if (set == null)
            {
                throw new KeyNotFoundException("プロファイルが見つかりません: " + name);
            }

            return CreateProfile(set);
        }

        /// <inheritdoc />
        public IEnumerable<string> ListNames()
        {
            return LoadMappingSets()
                .Where(mappingSet => mappingSet != null)
                .Select(mappingSet => mappingSet.name);
        }

        private List<MappingSet> LoadMappingSets()
        {
            using (var mappingsStream = File.OpenRead(_filePath))
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

        // 以下の JSON 読み込み用 DTO では、既定の名前照合で既存 JSON 形式を維持するため、
        // 小文字のプロパティ名を使用する

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
