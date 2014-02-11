using System.Collections.Generic;
using System.ServiceModel;

namespace ExportOffice.Web
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IService1" in both code and config file together.
    [ServiceContract]
    public interface IService1
    {
        [OperationContract]
        void DoWork();

        [OperationContract]
        byte[] DoExportExcel(List<FactModel> facts, List<Columns> headersList);

        [OperationContract]
        bool DoUploadFile(byte[] buffer);

        [OperationContract]
        byte[] DoExportWord();
    }
}
