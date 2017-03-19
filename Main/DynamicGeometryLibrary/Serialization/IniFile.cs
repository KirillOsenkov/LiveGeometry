using System;
using System.Collections.Generic;
using System.Windows.Media;

namespace DynamicGeometry
{
    public class IniFile
    {
        public IniFile(string[] lines)
        {
            Sections = new List<Section>();
            ReadSections(lines);
        }

        void ReadSections(string[] lines)
        {
            int i = 0;
            int last = lines.Length - 1;
            Section currentSection = null;

            while (i <= last)
            {
                string currentLine = lines[i];
                int currentLineLength = currentLine.Length;
                if (currentLine[0] == '[')
                {
                    if (currentLineLength < 3)
                    {
                        throw new Exception("Incomplete ini section header: {0}"
                             + currentLine);
                    }
                    currentSection = new Section(currentLine.Substring(1, currentLineLength - 2));
                    Sections.Add(currentSection);
                }
                else
                {
                    int equals = currentLine.IndexOf('=');
                    if (equals < 1 || equals > currentLineLength - 1)
                    {
                        throw new Exception("Incorrect 'name=value' syntax: {0}" + currentLine);
                    }
                    string name = currentLine.Substring(0, equals);
                    string value = currentLine.Substring(equals + 1, currentLineLength - equals - 1);
                    currentSection.Entries.Add(name, value);
                }
                i++;
            }
        }

        public List<Section> Sections { get; private set; }

        public class Section
        {
            public Section(string title)
            {
                Entries = new Dictionary<string, string>();
                Title = title;
            }

            public int GetTitleNumber(string prefix)
            {
                return int.Parse(Title.Substring(prefix.Length));
            }

            public string this[string index]
            {
                get
                {
                    string result = null;
                    Entries.TryGetValue(index, out result);
                    return result;
                }
            }

            public string ReadString(string key)
            {
                return this[key];
            }

            public double? TryReadDouble(string key)
            {
                var value = this[key];
                if (value == null)
                {
                    return null;
                }
                return double.Parse(value);
            }

            public double ReadDouble(string key)
            {
                return double.Parse(this[key]);
            }

            public string Title { get; set; }

            public Dictionary<string, string> Entries { get; private set; }

            public int ReadInt(string key)
            {
                var value = this[key];
                if (value == null)
                {
                    return 0;
                }

                return int.Parse(value);
            }

            public int? TryReadInt(string key)
            {
                var value = this[key];
                if (value == null)
                {
                    return null;
                }
                return int.Parse(value);
            }

            public Color ReadColor(string key)
            {
                if (this[key] == null)
                {
                    return Colors.Black;
                }

                var number = ReadInt(key);
                return number.ToColor();
            }

            public bool ReadBool(string key, bool defaultValue)
            {
                bool result = defaultValue;
                string value = this[key];
                if (value != null)
                {
                    bool.TryParse(value.ToLower(), out result);
                }
                return result;
            }
        }
    }
}
