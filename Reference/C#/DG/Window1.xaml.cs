using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Resources;
using System.Reflection;
using System.Threading;
using System.Collections;
using System.Diagnostics;

namespace DynamicGeometry
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        public Window1()
        {
            InitializeComponent();

            drawing = new Drawing(canvas1);
            Assembly asm = typeof(Window1).Assembly;
            string resourceName = asm.GetName().Name + ".g";
            ResourceManager rm = new ResourceManager(resourceName, asm);
            ResourceSet resourceSet = rm.GetResourceSet(Thread.CurrentThread.CurrentCulture, true, true);
            List<string> resources = new List<string>();
            foreach (DictionaryEntry resource in resourceSet)
            {
                Debug.WriteLine(resource.Key);
            }
            rm.ReleaseAllResources();

            toolBar1.ItemsSource = Behavior.Singletons;
        }

        Drawing drawing;

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Behavior b = (sender as Button).Tag as Behavior;
            if (b != null)
            {
                Drawing.Current.Behavior = b;
            }
        }
    }
}
