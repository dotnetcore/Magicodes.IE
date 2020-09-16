using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Tests.Models.Import;
using Microsoft.Extensions.DependencyInjection;
using Shouldly;
using System.Threading.Tasks;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class AppDependencyResolver_Tests
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        [Fact()]
        public void AppDependencyResolverNotInit_Test()
        {
            Assert.ThrowsAny<System.Exception>(() =>
            {
                AppDependencyResolver.Current.GetService<IImportResultFilter>();
            });
        }

        [Fact()]
        public void AppDependencyResolverGetService_Test()
        {
            //初始化容器
            var services = new ServiceCollection();
            //添加注入关系
            services.AddSingleton<IImportResultFilter, ImportResultFilterTest>();
            var serviceProvider = services.BuildServiceProvider();
            AppDependencyResolver.Init(serviceProvider);
            AppDependencyResolver.Current.GetService<IImportResultFilter>().ShouldNotBeNull();
            AppDependencyResolver.Dispose();
        }
    }
}
