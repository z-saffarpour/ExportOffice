using System;
using System.Runtime.Serialization;

namespace ExportOffice.Web
{
    [DataContract]
    public class Columns
    {
        [DataMember]
        public string Header { get; set; }
        [DataMember]
        public string ColumnType { get; set; }
    }
}