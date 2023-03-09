/*
 * Applica, Inc.
 * By José O. Lara
 * 2023.03.09
 */

namespace wpfApp_Demo.src.core.data.configuration {
    internal class DbfConfiguration {

        public string MyDbfPath { get; }

        public DbfConfiguration(string myDbfPath) => MyDbfPath = myDbfPath;

    }
}
