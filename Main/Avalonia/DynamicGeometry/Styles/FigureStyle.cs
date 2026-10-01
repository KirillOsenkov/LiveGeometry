using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    [PropertyGridName("Edit style")]
    public abstract partial class FigureStyle : IFigureStyle, IThemeOverridable, IConditionalProperties
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

        #region Themes

        // A style has its values, which are how it looks under the base theme (Light), and
        // for another theme may hold other values for some of its properties (a Dark override:
        // a darker fill, a lighter ink). Whoever draws with the style resolves it first
        // (Resolve): the style itself when the theme on screen has no override, else a copy
        // with the override's values set. The default styles of a drawing take their colors
        // from the theme's paper group (BindToTheme): the base value from the Light theme, an
        // override from every other, kept current when a theme color is tweaked.

        /// <summary>By theme name, the properties that differ under that theme, with their values</summary>
        [Ignore]
        public Dictionary<string, Dictionary<string, object>> Overrides { get; private set; } = new Dictionary<string, Dictionary<string, object>>();

        Dictionary<string, Func<AppTheme, object>> themeBindings;

        /// <summary>Under the named theme, the property has this value</summary>
        public void SetOverride(string theme, string property, object value)
        {
            if (!Overrides.TryGetValue(theme, out var values))
            {
                values = new Dictionary<string, object>();
                Overrides[theme] = values;
            }

            values[property] = value;
            OnPropertyChanged(property);
        }

        public void RemoveOverride(string theme, string property)
        {
            if (Overrides.TryGetValue(theme, out var values) && values.Remove(property))
            {
                if (values.Count == 0)
                {
                    Overrides.Remove(theme);
                }

                OnPropertyChanged(property);
            }
        }

        public void ClearOverrides(string theme)
        {
            if (Overrides.Remove(theme, out var values))
            {
                foreach (var property in values.Keys)
                {
                    OnPropertyChanged(property);
                }
            }
        }

        /// <summary>
        /// The button that drops the theme's own values (shown under a theme that has some):
        /// the style looks as it does under the base theme again. One undo step, which puts
        /// the values back.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridIcon(PropertyGridIcon.Cross)]
        public void SameAsBaseTheme()
        {
            string theme = AppTheme.Current.Name;
            if (!Overrides.TryGetValue(theme, out var values))
            {
                return;
            }

            var saved = new Dictionary<string, object>(values);
            var action = new CallMethodAction(
                () => ClearOverrides(theme),
                () =>
                {
                    foreach (var pair in saved)
                    {
                        SetOverride(theme, pair.Key, pair.Value);
                    }
                });
            var actionManager = StyleManager?.Drawing?.ActionManager;
            if (actionManager != null)
            {
                actionManager.RecordAction(action);
            }
            else
            {
                action.Execute();
            }
        }

        public virtual bool CanEdit(string propertyName)
        {
            if (propertyName == nameof(SameAsBaseTheme))
            {
                var theme = AppTheme.Current;
                return !AppTheme.IsBase(theme) && Overrides.TryGetValue(theme.Name, out var values) && values.Count > 0;
            }

            return true;
        }

        public virtual string Caption(string propertyName, string defaultCaption)
        {
            return propertyName == nameof(SameAsBaseTheme) ? "Same as in " + AppTheme.Base.Name : defaultCaption;
        }

        /// <summary>The property takes its value from every theme's colors: a default style's ink or fill</summary>
        public void BindToTheme(string property, Func<AppTheme, object> value)
        {
            themeBindings ??= new Dictionary<string, Func<AppTheme, object>>();
            themeBindings[property] = value;
            ReadFromTheme(property, value);
        }

        /// <summary>Reads the bound properties from the themes again (a theme color was tweaked)</summary>
        public void RefreshFromTheme()
        {
            if (themeBindings == null)
            {
                return;
            }

            foreach (var binding in themeBindings)
            {
                ReadFromTheme(binding.Key, binding.Value);
            }
        }

        void ReadFromTheme(string property, Func<AppTheme, object> value)
        {
            foreach (var theme in AppTheme.All)
            {
                var themeValue = value(theme);
                if (theme == AppTheme.Light)
                {
                    // the setter raises PropertyChanged, and the figures repaint
                    GetType().GetProperty(property).SetValue(this, themeValue);
                }
                else
                {
                    SetOverride(theme.Name, property, themeValue);
                }
            }
        }

        /// <summary>
        /// The style looks under every theme as it does under the base one and no longer
        /// follows the themes' colors: for a drawing whose paper stays light under every theme
        /// (a GeoGebra worksheet). Each override says the base value, so that a file keeps the
        /// style that way (<see cref="StyleManager.AddWithDefaults"/> takes a default without
        /// overrides for the theme's own).
        /// </summary>
        public void KeepBaseLook()
        {
            themeBindings = null;
            var type = GetType();
            foreach (var theme in Overrides.Keys.ToArray())
            {
                foreach (var property in Overrides[theme].Keys.ToArray())
                {
                    SetOverride(theme, property, type.GetProperty(property).GetValue(this));
                }
            }
        }

        public IFigureStyle Resolve()
        {
            return Resolve(AppTheme.Current.Name);
        }

        /// <summary>The style as it looks under the theme: itself without an override for it</summary>
        public IFigureStyle Resolve(string theme)
        {
            if (!Overrides.TryGetValue(theme, out var values) || values.Count == 0)
            {
                return this;
            }

            var result = (FigureStyle)MemberwiseClone();
            result.PropertyChanged = null; // the copy's setters must not repaint the original's figures
            result.Overrides = new Dictionary<string, Dictionary<string, object>>();
            result.themeBindings = null;
            var type = GetType();
            foreach (var pair in values)
            {
                type.GetProperty(pair.Key).SetValue(result, pair.Value);
            }

            return result;
        }

        #endregion

        public Style GetWpfStyle(IFigure figure)
        {
            Style result = new Style(typeof(FrameworkElement));
            ((FigureStyle)Resolve()).ApplyToWpfStyle(result, figure);
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

        /// <summary>A copy with the same values and overrides, but no name and no ties to the theme: it is the user's own now</summary>
        public virtual IFigureStyle Clone()
        {
            var result = (FigureStyle)this.MemberwiseClone();
            result.PropertyChanged = null;
            result.Name = "";
            result.themeBindings = null;
            result.Overrides = new Dictionary<string, Dictionary<string, object>>();
            foreach (var pair in Overrides)
            {
                result.Overrides[pair.Key] = new Dictionary<string, object>(pair.Value);
            }

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

        /// <summary>The values, then each theme's overrides: two styles that look the same under every theme have the same signature</summary>
        public virtual string GetSignature()
        {
            var result = GetBaseSignature();
            foreach (var theme in Overrides.Keys.OrderBy(name => name))
            {
                foreach (var value in OverrideValues(theme))
                {
                    result += " " + theme + "." + value.Name + "=" + SerializationService.Instance.Write(value);
                }
            }

            return result;
        }

        /// <summary>The values under the base theme alone: how the style looks in Light</summary>
        public string GetBaseSignature()
        {
            var values = IncludeByDefaultValueDiscoveryStrategy.Instance
                    .GetValues(this)
                    .Where(v => v.Name != "Name")
                    .Select(v => SerializationService.Instance.Write(v)?.ToString()); // a null string (no character) writes nothing
            return string.Join(" ", values.ToArray());
        }

        /// <summary>The theme's overrides as values, in the order of the properties</summary>
        public IEnumerable<ThemedValue> OverrideValues(string theme)
        {
            if (!Overrides.TryGetValue(theme, out var values))
            {
                yield break;
            }

            foreach (var property in IncludeByDefaultValueDiscoveryStrategy.Instance.GetValues(this))
            {
                if (values.ContainsKey(property.Name))
                {
                    yield return new ThemedValue(property, this, theme);
                }
            }
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

        // Below is code necessary to implement an "OK" button that displays in the property grid when editing a style.

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
        [PropertyGridName("OK")]
        [PropertyGridIcon(PropertyGridIcon.Check)]
        public void FinishEditing()
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
        [PropertyGridName("Delete this style")]
        [PropertyGridDestructive]
        public void Delete()
        {
            StyleManager.Remove(this);
            FinishEditing();
        }

#endif

    }
}