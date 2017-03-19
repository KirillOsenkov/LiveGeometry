using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Reflection;

namespace DynamicGeometry
{
    public class MEFHost
    {
        private MEFHost()
        {
            assemblies.Add(typeof(MEFHost).Assembly);
            Compose();
        }

        private AggregateCatalog aggregateCatalog;
        private CompositionContainer compositionContainer;

        #region Services

        [Import]
        public ISerializationService SerializationService { get; set; }

        [Import]
        public ICompilerService CompilerService { get; set; }

        #endregion

        private void Compose()
        {
            aggregateCatalog = new AggregateCatalog();
            foreach (var assembly in assemblies)
            {
                AddAssemblyToCatalog(assembly);
            }

            compositionContainer = new CompositionContainer(aggregateCatalog);
            compositionContainer.ComposeParts(this);
        }

        private void AddAssemblyToCatalog(Assembly assembly)
        {
            aggregateCatalog.Catalogs.Add(new AssemblyCatalog(assembly));
        }

        private List<Assembly> assemblies = new List<Assembly>();

        public IEnumerable<Assembly> Assemblies
        {
            get
            {
                return assemblies;
            }
        }

        private static MEFHost instance;
        public static MEFHost Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new MEFHost();
                }

                return instance;
            }
        }

        public void SatisfyImportsOnce(object instance)
        {
            compositionContainer.SatisfyImportsOnce(instance);
        }

        public void RegisterExtensionAssemblyFromType<T>()
        {
            var assembly = typeof(T).Assembly;
            assemblies.Add(assembly);
            if (aggregateCatalog != null)
            {
                AddAssemblyToCatalog(assembly);
            }
        }
    }
}
