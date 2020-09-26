#if NETSTANDARD
using Microsoft.Extensions.DependencyInjection;
using System;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    /// 
    /// </summary>
    public class AppDependencyResolver
    {
        private static AppDependencyResolver _resolver;

        /// <summary>
        /// 是否已初始化
        /// </summary>
        public static bool HasInit => _resolver != null;

        /// <summary>
        /// 
        /// </summary>
        public static AppDependencyResolver Current
        {
            get
            {
                if (_resolver == null)
                {
                    throw new Exception("AppDependencyResolver not initialized. You should initialize it in Startup class");
                }

                return _resolver;
            }
        }

        /// <summary>
        /// 初始化
        /// </summary>
        /// <example>
        /// public void Configure(IApplicationBuilder app, IHostingEnvironment env, ILoggerFactory loggerFactory)
        ///{
        ///    AppDependencyResolver.Init(app.ApplicationServices);
        ///    //all other code
        ///}
        /// </example>
        /// <param name="services"></param>
        public static void Init(IServiceProvider services)
        {
            _resolver = new AppDependencyResolver(services);
        }

        public void Dispose()
        {
            _serviceProvider = null;
            _resolver = null;
        }

        private IServiceProvider _serviceProvider;

        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="serviceType"></param>
        ///// <returns></returns>
        //public object GetService(Type serviceType)
        //{
        //    return _serviceProvider.GetService(serviceType);
        //}

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public T GetService<T>()
        {
            return _serviceProvider.GetService<T>();
        }

        private AppDependencyResolver(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }
    }
}
#endif 
