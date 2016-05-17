using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExceLintUI
{
    public struct CellColor
    {
        public int ColorIndex;
        public double Color;
    }

    public class ColorDict
    {
        readonly CellColor transparent = new CellColor { ColorIndex = 0, Color = 0 };

        private Dictionary<AST.Address, CellColor> _d = new Dictionary<AST.Address, CellColor>();

        public CellColor restoreColorAt(AST.Address addr)
        {
            if (_d.ContainsKey(addr))
            {
                return _d[addr];
            } else
            {
                return transparent;
            }
        }

        public void saveColorAt(AST.Address addr, CellColor c)
        {
            if (c.ColorIndex != 0)
            {
                if (_d.ContainsKey(addr))
                {
                    _d[addr] = c;
                } else
                {
                    _d.Add(addr, c);
                }
            }
        }

        public void Clear()
        {
            _d.Clear();
        }

        public Dictionary<AST.Address, CellColor> all()
        {
            return _d;
        }
    }
}
