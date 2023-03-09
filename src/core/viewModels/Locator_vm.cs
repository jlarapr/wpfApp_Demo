/*
 * Applica, Inc.
 * By José O. Lara
 * 2023.03.09
 */

namespace wpfApp_Demo.src.core.viewModels {
    using Microsoft.Extensions.DependencyInjection;
    public class LocatorVm {
        public static MainVm MainVm => App.ServiceProvider!.GetRequiredService<MainVm>();
    }
}
