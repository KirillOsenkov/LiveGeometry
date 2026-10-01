using System;
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
            }
            catch (Exception ex)
            {
                Drawing.RaiseError(this, ex);
            }
            finally
            {
                ForgetLoading();
            }
            PointBase.SuppressAutoLabelPoints = false;
        }

        /// <summary>
        /// Reading a file is not something to undo: the history starts empty, also when the
        /// file failed half way - or Undo would take the figures that did load away one by one
        /// </summary>
        void ForgetLoading()
        {
            var manager = Drawing.ActionManager;
            if (manager.RecordingTransaction != null)
            {
                manager.TransactionStack.Clear();
            }

            manager.Clear();
        }

        public void LoadDrawingFromDGF(string[] lines, string fileName)
        {
            try
            {
                ShowOperationDuration(() =>
                {
                    Clear();
                    Drawing.AddFromDGF(lines);
                    Drawing.Name = fileName;
                });
            }
            catch (Exception ex)
            {
                Drawing.RaiseError(this, ex);
            }
            finally
            {
                ForgetLoading();
            }
        }

        public void LoadDrawingFromDGF(string[] lines)
        {
            LoadDrawingFromDGF(lines, "");
        }

        /// <summary>A GeoGebra worksheet; points get labels only where the file shows them</summary>
        public void LoadDrawingFromGeoGebra(XElement worksheet, string fileName)
        {
            PointBase.SuppressAutoLabelPoints = true;
            try
            {
                Clear();
                Drawing.AddFromGeoGebra(worksheet);
                Drawing.Name = fileName;
            }
            catch (Exception ex)
            {
                Drawing.RaiseError(this, ex);
            }
            finally
            {
                ForgetLoading();
            }

            PointBase.SuppressAutoLabelPoints = false;
        }

        public void ShowOperationDuration(Action code)
        {
            var duration = Utilities.ElapsedTime(code);
            Drawing.RaiseStatusNotification(string.Format("Processed in {0} milliseconds", duration));
        }
    }
}
