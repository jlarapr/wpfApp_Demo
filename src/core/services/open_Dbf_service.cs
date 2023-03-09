/*
 * Applica, Inc.
 * By José O. Lara
 * 2023.03.09
 */

namespace wpfApp_Demo.src.core.services {
    using wpfApp_Demo.src.core.data.configuration;
    using wpfApp_Demo.src.core.data.repository;
    using wpfApp_Demo.src.core.interfaces.repository;
    using wpfApp_Demo.src.core.interfaces.services;
    using System.Data;
    using System.Threading.Tasks;

    internal class OpenDbfService : IOpenDbfService {
        private readonly DbfConfiguration MyDbfConfiguration;
        private readonly IOpenDbfRepository MyOpenDBFRepository;

        public OpenDbfService(DbfConfiguration myDbfConfiguration) {
            MyDbfConfiguration = myDbfConfiguration;
            MyOpenDBFRepository = new OpenDbfRepository(MyDbfConfiguration.MyDbfPath);
        }

        public Task<DataTable> GetAllAsDataTableAsync(string dbfPath) {
            return MyOpenDBFRepository.GetAllAsDataTableAsync(dbfPath);
        }

        public Task<DataTable> GetAllAsDataTableAsync() {
            return MyOpenDBFRepository.GetAllAsDataTableAsync();
        }

    }
}
