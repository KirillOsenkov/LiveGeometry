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
            var xml = ParseDrawing(drawingXml, out string problem);
            if (xml == null)
            {
                Drawing.RaiseStatusNotification(problem);
                return;
            }

            LoadDrawing(xml, fileName);
        }

        /// <summary>
        /// The text of a drawing as XML; null, with the reason in words, for a text that is
        /// none: a picture or a document picked by mistake, a file cut short, XML of another
        /// kind (which used to open as an empty drawing under the file's name, for Save to
        /// write over the file).
        /// </summary>
        public static XElement ParseDrawing(string text, out string problem)
        {
            problem = null;

            // checked, not tried: an exception, even caught, is an error report on screen
            if (string.IsNullOrWhiteSpace(text) || !text.TrimStart().StartsWith("<"))
            {
                problem = "This file is not a drawing.";
                return null;
            }

            XElement xml;
            try
            {
                xml = XElement.Parse(text);
            }
            catch (System.Xml.XmlException ex)
            {
                problem = "This file is damaged: " + ex.Message;
                return null;
            }

            if (xml.Name.LocalName != "Drawing")
            {
                problem = "This file is not a drawing.";
                return null;
            }

            return xml;
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

                // only a file read in full gives the drawing its name (and with it the
                // right to be saved over that file); otherwise the status says what is missing
                if (Drawing.LoadErrors == null)
                {
                    Drawing.Name = fileName;
                    Drawing.ClearStatus();
                }
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
