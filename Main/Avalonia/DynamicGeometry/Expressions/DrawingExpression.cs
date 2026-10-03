using System;
using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry
{
    /// <summary>
    /// A figure given by expressions (a point by coordinates, a circle or a line by
    /// equation): it depends on the figures they name, and on nothing else
    /// </summary>
    public interface IExpressionOwner
    {
        /// <summary>Always in the same order (X, then Y), which is the order the dependencies are listed in</summary>
        IEnumerable<DrawingExpression> Expressions { get; }
    }

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

            mValue = result.Expression;
            Dependencies = result.Dependencies;

            // The figure depends on what its expressions name, all of them, listed in the
            // order of the expressions (X's, then Y's). Not "this one's old ones out, its new
            // ones in at the end": a figure another expression names too went out with them
            // (X = A.X, Y = A.Y, then X edited: the point no longer followed A, nor went
            // with it), and undo of an edit left the list in another order than it was.
            var named = new List<IFigure>();
            var expressions = ParentFigure is IExpressionOwner owner ? owner.Expressions : new[] { this };
            var figures = ParentFigure.Drawing.Figures;

            // A figure being read is asked for its place before what it names is in the
            // drawing (an intersection works itself out as it is read): an expression that
            // names a figure doesn't compile yet, and the others, compiling, set the
            // dependencies to what they name - nothing. Its dependencies are then what the
            // file says until all of them compile. (A line by equation x = P.X lost P in a
            // file saved while the intersection on it was there, and no longer followed it.)
            if (!figures.Contains(ParentFigure)
                && expressions.Any(expression => expression != null
                    && expression != this
                    && expression.Dependencies == null
                    && !expression.Text.IsEmpty()))
            {
                return;
            }

            ParentFigure.UnregisterFromDependencies();
            foreach (var expression in expressions)
            {
                if (expression == null || expression.Dependencies == null)
                {
                    continue;
                }

                foreach (var dependency in expression.Dependencies)
                {
                    // another expression may still name a figure that has left: when a
                    // point is replaced, the expressions are compiled again one by one
                    if ((expression == this || figures.ContainsRecursively(dependency)) && !named.Contains(dependency))
                    {
                        named.Add(dependency);
                    }
                }
            }

            ParentFigure.Dependencies.SetItems(named);

            // Do the following only when the ParentFigure is already in Drawing.
            // RegisterWithDependencies gets called when the ParentFigure is added to Drawing.
            // Otherwise the dependency.Dependents will list the ParentFigure twice
            // and ultimately cause a consistency error.        - D.H.
            if (ParentFigure.Drawing.Figures.Contains(ParentFigure))
            {
                ParentFigure.RegisterWithDependencies();

                // a file being read is recalculated once all of it is in
                // (DrawingDeserializer.ReadDrawing), not after every expression it compiles
                if (!ParentFigure.Drawing.IsReading)
                {
                    ParentFigure.RecalculateAllDependents();
                }
            }
        }

        /// <summary>The text follows renamed figures; the compiled value holds the figures already</summary>
        public void RenameInExpression(ExpressionRenamer renamer)
        {
            Text = renamer.Rewrite(Text, isFunction: false);
        }

        /// <summary>Compiles the text again (<see cref="IRenamableExpressions.RebindExpressions"/>)</summary>
        public void Rebind()
        {
            if (!Text.IsEmpty())
            {
                Recalculate();
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
