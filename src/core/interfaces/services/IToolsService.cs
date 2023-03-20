/*
 * Applica, Inc.
 * By José O. Lara
 * 2023.03.20
 */


namespace wpfApp_Demo.src.core.interfaces.services {
    
    using System.Data;
    

    internal interface IToolsService {
        void TotalesToExcel(string excelFileName, string titel, string applicationName, string myDate, DataTable table);
    }
}
