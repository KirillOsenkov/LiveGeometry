using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Windows;

namespace DynamicGeometry
{
    public class Drawing
    {
        public Drawing(Canvas newParent)
        {
            if (Current == null)
            {
                Current = this;
            }
            Figures = new FigureGroup(this);
            Parent = newParent;
            Behavior = new Dragger();
        }

        private static Drawing mCurrent;
        public static Drawing Current
        {
            get
            {
                return mCurrent;
            }
            set
            {
                mCurrent = value;
            }
        }

        private Canvas mParent;
        public Canvas Parent
        {
            get
            {
                return mParent;
            }
            set
            {
                mParent = value;
                if (Behavior != null)
                {
                    Behavior.Drawing = this;
                }
            }
        }

        private Behavior mBehavior;
        public Behavior Behavior
        {
            get
            {
                return mBehavior;
            }
            set
            {
                if (mBehavior == value)
                {
                    return;
                }
                if (mBehavior != null)
                {
                    mBehavior.Drawing = null;
                }
                mBehavior = value;
                if (mBehavior != null)
                {
                    mBehavior.Drawing = this;
                }
            }
        }

        public FigureList Figures { get; set; }
    }
}
