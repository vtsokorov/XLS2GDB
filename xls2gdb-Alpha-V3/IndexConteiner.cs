using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xls2gdb
{
    class IndexConteiner
    {
        private List<Dictionary<int, int>> list = new List<Dictionary<int, int>>();

        public void add(Dictionary<int, int> dic)
        {
            list.Add(dic);
        }
        public void add(int[] left, int[] right)
        {
            if (left.Length == right.Length)
            {
                Dictionary<int, int> dic = new Dictionary<int, int>();
                for (int i = 0; i < left.Length; ++i)
                    dic.Add(left[i], right[i]);
                list.Add(dic);
            }
        }
        public void add(List<int> left, List<int> right)
        {
            if (left.Count == right.Count)
            {
                Dictionary<int, int> dic = new Dictionary<int, int>();
                for (int i = 0; i < left.Count; ++i)
                    dic.Add(left[i], right[i]);
                list.Add(dic);
            }
        }
        public void delete(int index)
        {
            list.RemoveAt(index);
        }
        public int size()
        {
            return list.Count;
        }
        public List<int> leftList(int index)
        { return list[index].Keys.ToList(); }
        public List<int> rightList(int index)
        { return list[index].Values.ToList(); }
    }
}
