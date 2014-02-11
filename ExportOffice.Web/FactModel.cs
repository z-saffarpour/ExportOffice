using System.Runtime.Serialization;

namespace ExportOffice.Web
{
    [DataContract]
    public class FactModel
    {

        private string _month;
        [DataMember]
        public string Month
        {
            get { return _month; }
            set
            {
                if (_month != value)
                {
                    _month = value;
                }
            }
        }

        private decimal _cost;
        [DataMember]
        public decimal Cost
        {
            get { return _cost; }
            set
            {
                if (_cost != value)
                {
                    _cost = value;
                }
            }
        }
    }
}
