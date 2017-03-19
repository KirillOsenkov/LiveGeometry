using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Windows;
using System.Collections.ObjectModel;
using System.Windows.Media;

namespace DynamicGeometry
{
    public abstract class Behavior
    {
        private static ReadOnlyCollection<Behavior> mSingletons;
        public static ReadOnlyCollection<Behavior> Singletons
        {
            get
            {
                if (mSingletons == null)
                {
                    mSingletons = InitializeBehaviors();
                }
                return mSingletons;
            }
        }

        private static ReadOnlyCollection<Behavior> InitializeBehaviors()
        {
            List<Behavior> result = new List<Behavior>();
            Type basic = typeof(Behavior);

            foreach (Type t in basic.Assembly.GetTypes())
            {
                if (basic.IsAssignableFrom(t) && !t.IsAbstract)
                {
                    Behavior instance = Activator.CreateInstance(t) as Behavior;
                    result.Add(instance);
                }
            }

            return result.AsReadOnly();
        }

        public virtual string Icon
        {
            get
            {
                return null;
            }
        }

        private Drawing mDrawing;
        public Drawing Drawing
        {
            get
            {
                return mDrawing;
            }
            set
            {
                if (mDrawing != null)
                {
                    Parent = null;
                }
                mDrawing = value;
                if (mDrawing != null)
                {
                    Parent = mDrawing.Parent;
                }
            }
        }

        protected virtual void AbortAndSetDefaultTool()
        {
            Reset();
            Drawing.Behavior = Behavior.Singletons.FirstOrDefault(b => b is Dragger);
        }

        protected virtual void Reset()
        {

        }

        private Canvas mParent;
        protected Canvas Parent
        {
            get
            {
                return mParent;
            }
            set
            {
                if (mParent != null)
                {
                    mParent.MouseDown -= MouseDown;
                    mParent.MouseMove -= MouseMove;
                    mParent.MouseUp -= MouseUp;
                }
                mParent = value;
                if (mParent != null)
                {
                    mParent.MouseDown += MouseDown;
                    mParent.MouseMove += MouseMove;
                    mParent.MouseUp += MouseUp;
                }
            }
        }

        protected virtual void MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.RightButton == System.Windows.Input.MouseButtonState.Pressed)
            {
                AbortAndSetDefaultTool();
            }
        }

        protected virtual void MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            
        }

        protected virtual void MouseUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
        }

        protected virtual FreePoint CreatePointAtCurrentPosition(System.Windows.Input.MouseButtonEventArgs e)
        {
            FreePoint p = Factory.CreateFreePoint();
            Drawing.Figures.Add(p);
            p.MoveTo(Coordinates(e));
            return p;
        }

        protected System.Windows.Point Coordinates(System.Windows.Input.MouseEventArgs e)
        {
            return e.GetPosition(Parent);
        }
    }
}
