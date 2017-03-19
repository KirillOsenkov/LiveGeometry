using System.Collections.Generic;
using System.Linq;
using System.Windows;
using DynamicGeometry;
using GuiLabs.Undo;

namespace LiveGeometry
{
    public partial class Page
    {
        Settings PageSettings;

        public class Settings
        {
            public Settings(Page page)
            {
                Page = page;
                this.DebugInfo = new DebugInformation(page);

                // create Polar List
                DynamicGeometry.Settings.Instance.PolarItems = new[]
                {
                    "5",
                    "10",
                    "15",
                    "18",
                    "22.5",
                    "30",
                    "45",
                    "90"               
                };

                PolarValue a = new PolarValue();
                a.Val = 30;
                DynamicGeometry.Settings.Instance.PolarIncrement = a;                
            }

            Page Page;

            public bool ShowToolbar
            {
                get
                {
                    return Page.ShowToolbar;
                }
                set
                {
                    Page.ShowToolbar = value;
                }
            }

            [PropertyGridVisible]
            [PropertyGridName("Fullscreen (F9)")]
            public void Fullscreen()
            {
                Application.Current.Host.Content.IsFullScreen =
                    !Application.Current.Host.Content.IsFullScreen;
            }

            [PropertyGridVisible]
            [PropertyGridName("Show text in toolbar buttons")]
            public bool ShowToolbarText
            {
                get
                {
                    return Page.ShowToolbarButtonText;
                }
                set
                {
                    Page.ShowToolbarButtonText = value;
                }
            }

            [PropertyGridVisible]
            [PropertyGridName("Use these letters for point names")]
            public string PointAlphabet
            {
                get
                {
                    return DynamicGeometry.Settings.Instance.PointAlphabet;
                }
                set
                {
                    DynamicGeometry.Settings.Instance.PointAlphabet = value;
                }
            }

            [PropertyGridVisible]
            [PropertyGridName("Cursor Tolerance")]
            [Domain(1, 20)]
            public double CursorTolerance
            {
                get
                {
                    return DynamicGeometry.Settings.Instance.CursorTolerance;
                }
                set
                {
                    DynamicGeometry.Settings.Instance.CursorTolerance = value;
                }
            }

            public class DebugInformation
            {
                public DebugInformation(Page page)
                {
                    Page = page;
                }

                Page Page;

                [PropertyGridVisible]
                public Drawing Drawing
                {
                    get
                    {
                        return Page.drawingHost.CurrentDrawing;
                    }
                }

                [PropertyGridVisible]
                public string DrawingXml
                {
                    get
                    {
                        return Page.drawingHost.CurrentDrawing.SaveAsText();
                    }
                }

                [PropertyGridVisible]
                public IEnumerable<IAction> UndoBuffer
                {
                    get
                    {
                        return Page.drawingHost.CurrentDrawing.ActionManager.EnumUndoableActions().ToArray();
                    }
                }
            }
            
            [PropertyGridVisible]
            [PropertyGridName("Snap Grid Spacing")]
            public double SnapGridSpacing
            {
                get
                {
                    return DynamicGeometry.Settings.Instance.SnapGridSpacing;
                }
                set
                {
                    DynamicGeometry.Settings.Instance.SnapGridSpacing = value;
                }
            }

            [PropertyGridVisible]
            public DebugInformation DebugInfo { get; set; }


            //[PropertyGridVisible]
            //[PropertyGridName("Show text in toolbar buttons")]
            //public bool ShowToolbarText
            //{
            //    get
            //    {
            //        return Page.ShowToolbarButtonText;
            //    }
            //    set
            //    {
            //        Page.ShowToolbarButtonText = value;
            //    }
            //}

            //Polar settings
            [PropertyGridVisible]
            [PropertyGridName("Polar")]
            public PolarValue Polar
            {
                get
                {
                    return DynamicGeometry.Settings.Instance.PolarIncrement;
                }
                set
                {
                    DynamicGeometry.Settings.Instance.PolarIncrement = value;
                }
            }
            
            // Incerement values
            [PropertyGridVisible]
            [PropertyGridName("New Increment")]
            public string AdditinalPolarIncrement
            {
                get
                {
                    return DynamicGeometry.Settings.Instance.TempNewIncrement;
                }
                set
                {
                    DynamicGeometry.Settings.Instance.TempNewIncrement = value;
                }
            }

            // Add new Polar Increment
            [PropertyGridVisible]
            [PropertyGridName("Add Increment")]
            public void AddNewIncrement()
            {
                if (DynamicGeometry.Math.IsDoubleValid(DynamicGeometry.Settings.Instance.TempNewIncrement))
                {
                    string[] InitialItemList;
                    InitialItemList = (string[])DynamicGeometry.Settings.Instance.PolarItems;

                    int CurLen = InitialItemList.Length;

                    string[] ConvertedItemList = new string[CurLen + 1];

                    for (int i = 0; i < CurLen; i++)
                    {
                        ConvertedItemList[i] = InitialItemList[i];
                    }
                    ConvertedItemList[CurLen] = DynamicGeometry.Settings.Instance.TempNewIncrement;
                    DynamicGeometry.Settings.Instance.TempNewIncrement = "";

                    DynamicGeometry.Settings.Instance.PolarItems = ConvertedItemList;
                }
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            OpenSettings();
        }

        private void OpenSettings()
        {
            if (drawingHost.PropertyGrid.Selection == PageSettings)
            {
                drawingHost.ShowProperties(null);
            }
            else
            {
                drawingHost.ShowProperties(PageSettings);
            }
        }
    }
}
