using System.ComponentModel;

namespace ExportOffice.Model
{
    public class FactMode : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(name));
        }

        private int _cost;
        public int Cost
        {
            get { return _cost; }
            set
            {
                if (_cost != value)
                {
                    _cost = value;
                    OnPropertyChanged("Cost");
                }
            }
        }

        private string _month;
        public string Month
        {
            get{return _month;}
            set
            {
                if (_month != value)
                {
                    _month = value;
                    OnPropertyChanged("Month");
                }
            }
        }
    }
}
