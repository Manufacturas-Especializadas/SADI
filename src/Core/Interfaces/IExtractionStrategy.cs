using Core.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Core.Interfaces
{
    public interface IExtractionStrategy
    {
        string ClientFolderIdentifier { get; }

        string DocumentTypeSubFolder {  get; }

        List<PurchaseOrderItem> Extract(string filePath);
    }
}