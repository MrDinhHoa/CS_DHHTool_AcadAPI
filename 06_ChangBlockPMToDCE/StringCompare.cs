using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _06_ChangBlockPMToDCE
{
    public class NaturalStringComparer : IComparer<BlockRebarDCEInfor>, IComparer
    {
        public int Compare(BlockRebarDCEInfor x, BlockRebarDCEInfor y)
        {
            return Compare(x.SH as string, y.SH as string);
        }
        public int Compare(string x, string y)
        {
            if (x == null || y == null)
            {
                return Comparer<string>.Default.Compare(x, y);
            }

            int ix = 0, iy = 0;
            while (ix < x.Length && iy < y.Length)
            {
                if (char.IsDigit(x[ix]) && char.IsDigit(y[iy]))
                {
                    // Lấy toàn bộ phần số
                    string nx = GetNumber(x, ref ix);
                    string ny = GetNumber(y, ref iy);

                    int result = CompareNumbers(nx, ny);
                    if (result != 0)
                    {
                        return result;
                    }
                }
                else
                {
                    int result = x[ix].CompareTo(y[iy]);
                    if (result != 0)
                    {
                        return result;
                    }

                    ix++;
                    iy++;
                }
            }

            return x.Length.CompareTo(y.Length);
        }
        private string GetNumber(string s, ref int i)
        {
            int start = i;
            while (i < s.Length && char.IsDigit(s[i]))
            {
                i++;
            }
            return s.Substring(start, i - start);
        }
        private int CompareNumbers(string x, string y)
        {
            // So sánh hai chuỗi số
            if (x.Length != y.Length)
            {
                return x.Length.CompareTo(y.Length);
            }
            return string.Compare(x, y, StringComparison.Ordinal);
        }

        public int Compare(object x, object y)
        {
            throw new NotImplementedException();
        }
    }
}
