using System;
using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry
{
    public class DrawingExpression : IValueProvider
    {
        public DrawingExpression(IFigure parent)
        {
            ParentFigure = parent;
            IsValid = false;
        }

        public DrawingExpression(IFigure parent, string name, string expressionText)
        {
            ParentFigure = parent;
            Name = name;
            Text = expressionText;
            IsValid = false;
        }

        public IFigure ParentFigure { get; private set; }

        Func<double> mValue;
        public Func<double> Value
        {
            get
            {
                if (mValue == null)
                {
                    Recalculate();
                }
                return mValue;
            }
            set
            {
                mValue = value;
                IsValid = value == null;
            }
        }
        public string Text { get; set; }
        public bool IsValid { get; private set; }
        List<IFigure> Dependencies { get; set; }

        public void Recalculate()
        {
            var result = Compiler.Instance.CompileExpression(
                 ParentFigure.Drawing,
                 Text,
                 f => !f.DependsOn(ParentFigure));
            IsValid = result.IsSuccess;
            if (!IsValid)
            {
                return;
            }

            ParentFigure.UnregisterFromDependencies();

            mValue = result.Expression;

            if (!Dependencies.IsEmpty())
            {
                ParentFigure.Dependencies.RemoveAll(Dependencies);
            }
            Dependencies = result.Dependencies;
            if (ParentFigure.Dependencies != null)
            {
                ParentFigure.Dependencies.Merge(Dependencies);
            }
            else
            {
                ParentFigure.Dependencies.SetItems(Dependencies);
            }

            // Do the following only when the ParentFigure is already in Drawing.
            // RegisterWithDependencies gets called when the ParentFigure is added to Drawing.
            // Otherwise the dependency.Dependents will list the ParentFigure twice
            // and ultimately cause a consistency error.        - D.H.
            if (ParentFigure.Drawing.Figures.Contains(ParentFigure))
            {
                ParentFigure.RegisterWithDependencies();
                ParentFigure.RecalculateAllDependents();
            }
        }

        public override string ToString()
        {
            return Text;
        }

        public event Action ValueChanged;

        public void RaiseValueChanged()
        {
            if (ValueChanged != null)
            {
                ValueChanged();
            }
        }

        public T GetValue<T>()
        {
            return (T)(object)Text;
        }

        public bool CanSetValue
        {
            get { return true; }
        }

        public void SetValue<T>(T value)
        {
            Text = value.ToString();
            Recalculate();
            RaiseValueChanged();
        }

        public object Parent
        {
            get { return this; }
        }

        public Type Type
        {
            get { return typeof(DrawingExpression); }
        }

        public string Name { get; set; }

        public string DisplayName
        {
            get { return Name; }
        }

        public T GetAttribute<T>() where T : Attribute
        {
            return null;
        }

        public System.Collections.Generic.IEnumerable<T> GetAttributes<T>() where T : Attribute
        {
            return null;
        }

        public string GetSignature()
        {
            return Name;
        }
    }
}
