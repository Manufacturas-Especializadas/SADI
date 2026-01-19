using Core.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Core.Interfaces
{
    public interface IReportGenerator
    {
        //void GenerateExcelReport(List<PurchaseOrderItem> data, string outputFilePath);

        void AppendToMasterLog(List<PurchaseOrderItem> data, string masterFilePath);

        HashSet<string> GetProcessFileNames(string masterFilePath);
    }
}