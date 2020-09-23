using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public abstract partial class FigureStyle : IFigureStyle
    {
        string name = "";
        //[PropertyGridVisible]
        public string Name 
        {
            get
            {
                return name;
            }
            set
            {
                // Prevent invalid (duplicate) style names.  Stylemanager is not designed to handle duplicate names.
                if (StyleManager != null)
                {
                    if (StyleManager.NameIsValid(value))
                    {
                        name = value;
                    }
                    else
                    {
                        MessageBox.Show("There is already a style with this name");
                    }
                }
                else
                {
                    name = value;
                }
#if !PLAYER
                if (CurrentEditInfo.PropertyGrid != null)
                {
                    CurrentEditInfo.PropertyGrid.UpdateHeader();
                }
#endif
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [Ignore]
        public StyleManager StyleManager { get; set; }

        public Style GetWpfStyle(IFigure figure)
        {
            Style result = new Style(typeof(FrameworkElement));
            ApplyToWpfStyle(result, figure);
            return result;
        }

        protected virtual void ApplyToWpfStyle(Style existingStyle, IFigure figure)
        {
            if (figure != null)
            {
                if (!figure.Enabled)
                {
                    existingStyle.Setters.Add(new Setter(FrameworkElement.OpacityProperty, 0.2));
                }
            }
        }

        public virtual IFigureStyle Clone()
        {
            var result = (IFigureStyle)this.MemberwiseClone();
            result.Name = "";
            return result;
        }

        public virtual IEnumerable<IFigureStyle> GetCompatibleStyles()
        {
            var result = StyleManager.GetCompatibleStyles(this.GetType());
            return result;
        }

        public abstract FrameworkElement GetSampleGlyph();

        protected virtual void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public virtual void OnApplied(IFigure figure, FrameworkElement element)
        {
        }

        public virtual string GetSignature()
        {
            var values = IncludeByDefaultValueDiscoveryStrategy.Instance
                    .GetValues(this)
                    .Where(v => v.Name != "Name")
                    .Select(v => SerializationService.Instance.Write(v).ToString());
            var result = string.Join(" ", values.ToArray());
            return result;
        }

#if !PLAYER
        public override string ToString()
        {
            var result = Name;
#if !TABULA
            result += "\n" + GetSignature();
#endif
            return result;
        }

        // Below is code necessary to implement a "Done" button that displays in the property grid when editing a style.

        EditInfo mCurrentEditInfo;
        [Ignore]
        public EditInfo CurrentEditInfo
        {
            get
            {
                if (mCurrentEditInfo == null)
                {
                    mCurrentEditInfo = new EditInfo();
                }
                return mCurrentEditInfo;
            }
            set
            {
                mCurrentEditInfo = value;
            }
        }

        public class EditInfo
        {
            public PropertyGrid PropertyGrid { get; set; }
            public object ParentObject { get; set; }
            public ActionManager ActionManager { get; set; }
        }

        [PropertyGridVisible]
        [PropertyGridName("Done")]
        public void DoneButton()
        {
            if (CurrentEditInfo.PropertyGrid != null && CurrentEditInfo.ParentObject != null)
            {
                CurrentEditInfo.PropertyGrid.Show(CurrentEditInfo.ParentObject, CurrentEditInfo.ActionManager);
                CurrentEditInfo.PropertyGrid = null;
                CurrentEditInfo.ParentObject = null;
                CurrentEditInfo.ActionManager = null;
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Delete This Style")]
        public void Delete()
        {
            StyleManager.Remove(this);
            DoneButton();
        }

#endif

    }
}