using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Castle.MicroKernel.Registration;
using Castle.Windsor;
using ExcelCustomTools.TextFixer;

namespace ExcelCustomTools.AddIn
{
    internal class DependencyContainer
    {
        readonly WindsorContainer _container = new WindsorContainer();

        public void Initialize()
        {
            this._container.Register(Component.For<IReplacementTableProvider>().ImplementedBy<ReplacementTableFileProvider>());
            this._container.Register(Component.For<ITextFixer>().ImplementedBy<TextFixer.TextFixer>());
        }

        public T Resolve<T>()
        {
            return this._container.Resolve<T>();
        }
    }
}
