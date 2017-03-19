using System;
using System.Collections.Generic;
using System.Linq;
using Expr = System.Linq.Expressions;
using System.Text;
using System.Windows;
using System.Linq.Expressions;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows.Controls;

namespace DynamicGeometry
{
    public interface IFigure
    {
        IFigureList Dependencies { get; }
        IFigureList Dependents { get; }
        void Recalculate();
        IFigure HitTest(Point point);
        void OnPlacingOnContainer(Canvas newContainer);
        void OnRemovingFromContainer(Canvas leavingContainer);
    }

    public interface IPoint : IFigure
    {
        Point Coordinates { get; }
    }

    public interface ILine : IFigure
    {
        PointPair Coordinates { get; }
    }

    public interface ICircle : IFigure
    {
        Point Center { get; }
        double Radius { get; }
    }

    public class ExpectedDependencyList : List<Type>
    {
        public ExpectedDependencyList()
        {

        }

        public static ExpectedDependencyList None = Create();
        public static ExpectedDependencyList PointPoint = Create<IPoint, IPoint>();

        public ExpectedDependencyList(params Type[] types)
        {
            AddRange(types);
        }

        public static ExpectedDependencyList Create()
        {
            return new ExpectedDependencyList();
        }

        public static ExpectedDependencyList Create<T>()
        {
            return new ExpectedDependencyList(typeof(T));
        }

        public static ExpectedDependencyList Create<T, T2>()
        {
            return new ExpectedDependencyList(typeof(T), typeof(T2));
        }

        public static ExpectedDependencyList Create<T, T2, T3>()
        {
            return new ExpectedDependencyList(typeof(T), typeof(T2), typeof(T3));
        }
    }

    public abstract class FigureBase : IFigure
    {
        public abstract ExpectedDependencyList RequiredDependencies { get; }

        private IFigureList m_Dependencies;
        public IFigureList Dependencies
        {
            get { return m_Dependencies; }
            protected set { m_Dependencies = value; }
        }

        private IFigureList m_Dependents = new FigureList();
        public IFigureList Dependents
        {
            get { return m_Dependents; }
            protected set { m_Dependents = value; }
        }

        protected void RegisterWithDependencies()
        {
            foreach (var dependency in Dependencies)
            {
                dependency.Dependents.Add(this);
            }
        }

        public Point Point(int index)
        {
            return (Dependencies[index] as IPoint).Coordinates;
        }

        public double Number(int index)
        {
            return (Dependencies[index] as INumber).Value;
        }

        public PointPair Line(int index)
        {
            return (Dependencies[index] as ILine).Coordinates;
        }

        public virtual void OnPlacingOnContainer(Canvas newContainer)
        {

        }

        public virtual void OnRemovingFromContainer(Canvas leavingContainer)
        {

        }

        public abstract void Recalculate();

        public abstract IFigure HitTest(Point point);
    }

    public interface INumber : IFigure
    {
        double Value { get; set; }
    }

    public struct PointPair
    {
        public Point P1;
        public Point P2;
    }

    public static class Extensions
    {
        public static Point Point(this IFigure figure, int index)
        {
            return (figure.Dependencies[index] as IPoint).Coordinates;
        }

        public static PointPair Line(this IFigure figure, int index)
        {
            return (figure.Dependencies[index] as ILine).Coordinates;
        }
    }
}
