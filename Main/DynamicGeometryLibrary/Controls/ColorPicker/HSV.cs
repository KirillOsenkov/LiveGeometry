using System;

namespace SilverlightContrib.Controls
{
    /// <summary>
    /// Data structure that represents a HSV value.
    /// </summary>
    internal struct HSV
    {
        private readonly double m_hue;
        private readonly double m_saturation;
        private readonly double m_value;

        public HSV(double hue, double saturation, double value)
        {
            m_hue = hue;
            m_saturation = saturation;
            m_value = value;
        }

        public double Hue
        {
            get { return m_hue; }
        }

        public double Saturation
        {
            get { return m_saturation; }
        }

        public double Value
        {
            get { return m_value; }
        }
    }
}
