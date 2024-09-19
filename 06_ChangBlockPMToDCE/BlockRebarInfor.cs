using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BIMSoftLib;
using BIMSoftLib.MVVM;

namespace _06_ChangBlockPMToDCE
{
    public class BlockRebarDCEInfor : PropertyChangedBase
    {
        private string _blockName;
        public string BlockName
        {
            get { return _blockName; }
            set { _blockName = value; OnPropertyChanged(nameof(BlockName)); }
        }


        private string _sH;
        public string SH
        {
            get { return _sH; }
            set { _sH = value; OnPropertyChanged(nameof(SH)); }
        }

        private string _dK;
        public string DK
        {
            get { return _dK; }
            set { _dK = value; OnPropertyChanged(nameof(DK)); }
        }


        private string _sL;
        public string SL
        {
            get { return _sL; }
            set { _sL = value; OnPropertyChanged(nameof(SL)); }
        }

        private string _sKC;
        public string SCK
        {
            get { return _sKC; }
            set { _sKC = value; OnPropertyChanged(nameof(SCK)); }
        }

        private string _m;
        public string M
        {
            get { return _m; }
            set { _m = value; OnPropertyChanged(nameof(M)); }
        }

        private string _m1;
        public string M1
        {
            get { return _m1; }
            set { _m1 = value; OnPropertyChanged(nameof(M1)); }
        }

        private string _m2;
        public string M2
        {
            get { return _m2; }
            set { _m2 = value; OnPropertyChanged(nameof(M2)); }
        }

        private string _l;
        public string L
        {
            get { return _l; }
            set { _l = value; OnPropertyChanged(nameof(L)); }
        }

        private string _l1;
        public string L1
        {
            get { return _l1; }
            set { _l1 = value; OnPropertyChanged(nameof(L1)); }
        }

        private string _l2;
        public string L2
        {
            get { return _l2; }
            set { _l2 = value; OnPropertyChanged(nameof(L2)); }
        }

        private string _l3;
        public string L3
        {
            get { return _l3; }
            set { _l3 = value; OnPropertyChanged(nameof(L3)); }
        }

        private string _l4;
        public string L4
        {
            get { return _l4; }
            set { _l4 = value; OnPropertyChanged(nameof(L4)); }
        }

        private string _l5;
        public string L5
        {
            get { return _l5; }
            set { _l5 = value; OnPropertyChanged(nameof(L5)); }
        }
    }
}
