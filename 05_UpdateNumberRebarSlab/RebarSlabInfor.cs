using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BIMSoftLib;

using BIMSoftLib.MVVM;

namespace _05_UpdateNumberRebarSlab
{
    public class RebarSlabInfor : PropertyChangedBase
    {
        private string _d2;
        public string D2
        {
            get { return _d2; }
            set { _d2 = value; OnPropertyChanged(nameof(D2)); }
        }

        private string _d1;
        public string D1
        {
            get { return _d1; }
            set { _d1 = value; OnPropertyChanged(nameof(D1)); }
        }

        private string _d3;
        public string D3
        {
            get { return _d3; }
            set { _d3 = value; OnPropertyChanged(nameof(D3)); }
        }

        private string _nO;
        public string NO
        {
            get { return _nO; }
            set { _nO = value; OnPropertyChanged(nameof(NO)); }
        }

        private string _nIE;
        public string NIE
        {
            get { return _nIE; }
            set { _nIE = value; OnPropertyChanged(nameof(NIE)); }
        }

        private string _dIA;
        public string DIA
        {
            get { return _dIA; }
            set { _dIA = value; OnPropertyChanged(nameof(DIA)); }
        }

        private string _qOE;
        public string QOE
        {
            get { return _qOE; }
            set { _qOE = value; OnPropertyChanged(nameof(QOE)); }
        }

        private string _lO;
        public string LO
        {
            get { return _lO; }
            set { _lO = value; OnPropertyChanged(nameof(LO)); }
        }

        private string _lA;
        public string LA
        {
            get { return _lA; }
            set { _lA = value; OnPropertyChanged(nameof(LA)); }
        }
    }
}
