using System.IO;
using System.IO.IsolatedStorage;
using System.Xml.Linq;
using DynamicGeometry;

namespace LiveGeometry
{
    public static class IsolatedStorage
    {
        const string wildcard = @"Macros\*.xml";

        public static void AddTool(XElement xml)
        {
            var fileNameAttribute = xml.Attribute("FileName");
            if (fileNameAttribute != null)
            {
                return;
            }
            var storage = IsolatedStorageFile.GetUserStoreForApplication();
            var fileName = GenerateNewFileName(storage);
            xml.Add(new XAttribute("FileName", fileName));
            using (var writer = new StreamWriter(new IsolatedStorageFileStream(fileName, FileMode.Create, storage)))
            {
                writer.WriteLine(xml.ToString());
            }
        }

        public static void SaveTool(XElement xml)
        {
            var fileName = xml.Attribute("FileName").Value;
            var storage = IsolatedStorageFile.GetUserStoreForApplication();
            using (var writer = new StreamWriter(new IsolatedStorageFileStream(fileName, FileMode.Create, storage)))
            {
                writer.WriteLine(xml.ToString());
            }
        }

        private static string GenerateNewFileName(IsolatedStorageFile storage)
        {
            string fileName = @"Macros\macro";
            string suffix = "";
            int suffixNumber = 0;
            while (storage.FileExists(fileName + suffix + ".xml"))
            {
                suffixNumber++;
                suffix = suffixNumber.ToString();
            }
            return fileName + suffix + ".xml";
        }

        public static void RemoveTool(XElement xml)
        {
            var storage = IsolatedStorageFile.GetUserStoreForApplication();
            var fileName = xml.Attribute("FileName").Value;
            if (storage.FileExists(fileName))
            {
                storage.DeleteFile(fileName);
            }
        }

        public static T GetSetting<T>(string settingName)
        {
            return GetSetting(settingName, default(T));
        }

        public static T GetSetting<T>(string settingName, T defaultValue)
        {
            T result = defaultValue;

            if (!IsolatedStorageSettings.ApplicationSettings.TryGetValue(settingName, out result))
            {
                result = defaultValue;
            }
            return result;
        }

        public static void LoadAllTools()
        {
            var storage = IsolatedStorageFile.GetUserStoreForApplication();
            if (!storage.DirectoryExists("Macros"))
            {
                storage.CreateDirectory("Macros");
            }
            var names = storage.GetFileNames(wildcard);
            foreach (var name in names)
            {
                using (var reader = new StreamReader(storage.OpenFile(@"Macros\" + name, FileMode.Open)))
                {
                    string tool = reader.ReadToEnd();
                    UserDefinedTool.AddFromString(tool);
                }
            }
        }

        public static void RegisterToolStorage()
        {
            ToolStorage.Instance = new IsolatedStorageBasedToolStorage();
        }

        public static void SaveSetting<T>(string settingName, T value)
        {
            IsolatedStorageSettings.ApplicationSettings[settingName] = value;
        }
    }

    public class IsolatedStorageBasedToolStorage : ToolStorage
    {
        public override void AddTool(UserDefinedTool newBehavior)
        {
            IsolatedStorage.AddTool(newBehavior.RootElement);
        }

        public override void RemoveTool(UserDefinedTool behavior)
        {
            IsolatedStorage.RemoveTool(behavior.RootElement);
        }

        public override void RenameTool(UserDefinedTool behavior, string newName)
        {
            IsolatedStorage.SaveTool(behavior.RootElement);
        }
    }

    public class IsolatedStorageBasedSettings : Settings
    {
        public override bool AutoLabelPoints
        {
            get
            {
                return IsolatedStorage.GetSetting<bool>("AutoLabelPoints");
            }
            set
            {
                IsolatedStorage.SaveSetting("AutoLabelPoints", value);
            }
        }

        public override bool ShowGrid
        {
            get
            {
                return IsolatedStorage.GetSetting<bool>("ShowGrid");
            }
            set
            {
                IsolatedStorage.SaveSetting("ShowGrid", value);
            }
        }

        public override double SnapGridSpacing
        {
            get
            {
                return IsolatedStorage.GetSetting<double>("SnapGridSpacing", 1);
            }
            set
            {
                IsolatedStorage.SaveSetting("SnapGridSpacing", value);
            }
        }

        public override bool ShowFigureExplorer
        {
            get
            {
                return IsolatedStorage.GetSetting<bool>("ShowFigureExplorer");
            }
            set
            {
                IsolatedStorage.SaveSetting("ShowFigureExplorer", value);
            }
        }

        public override string PointAlphabet
        {
            get
            {
                return IsolatedStorage.GetSetting<string>("PointAlphabet", "ABCDEFGHIJKLMNOPQRSTUVWXYZ");
            }
            set
            {
                IsolatedStorage.SaveSetting("PointAlphabet", value);
            }
        }

        public override double CursorTolerance
        {
            get
            {
                return IsolatedStorage.GetSetting<double>("Tolerance", 5);
            }
            set
            {
                IsolatedStorage.SaveSetting("Tolerance", value);
            }
        }
    }
}
