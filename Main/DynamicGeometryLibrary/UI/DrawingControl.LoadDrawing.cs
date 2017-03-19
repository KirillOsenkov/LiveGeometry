using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml.Linq;

namespace DynamicGeometry
{
    public partial class DrawingControl
    {
        private void mCurrentDrawing_DocumentOpenRequested(object sender, Drawing.DocumentOpenRequestedEventArgs e)
        {
            LoadDrawing(e.DocumentXml);
        }

        public void LoadDrawing(string drawingXml, string fileName)
        {
            XElement xml = null;
            try
            {
                xml = XElement.Parse(drawingXml);
            }
            catch (Exception ex)
            {
                Drawing.RaiseStatusNotification("Invalid file format: " + ex.ToString());
                return;
            }
            LoadDrawing(xml, fileName);
        }

        public void LoadDrawing(string drawingXml)
        {
            LoadDrawing(drawingXml, "");
        }

        public virtual void LoadDrawing(XElement element, string fileName)
        {
            PointBase.SuppressAutoLabelPoints = true;
            try
            {
                Clear();
                Drawing.AddFromXml(element);
                Drawing.Name = fileName;
                Drawing.ClearStatus();
                Drawing.ActionManager.Clear();
            }
            catch (Exception ex)
            {
                Drawing.RaiseError(this, ex);
            }
            PointBase.SuppressAutoLabelPoints = false;
        }

        public void LoadDrawingFromDGF(string[] lines, string fileName)
        {
            try
            {
                ShowOperationDuration(() =>
                {
                    Clear();
                    Drawing.AddFromDGF(lines);
                    Drawing.ActionManager.Clear();
                    Drawing.Name = fileName;
                });
            }
            catch (Exception ex)
            {
                Drawing.RaiseError(this, ex);
            }
        }

        public void LoadDrawingFromDGF(string[] lines)
        {
            LoadDrawingFromDGF(lines, "");
        }

        public void ShowOperationDuration(Action code)
        {
            var duration = Utilities.ElapsedTime(code);
            Drawing.RaiseStatusNotification(string.Format("Processed in {0} milliseconds", duration));
        }
    }
}
