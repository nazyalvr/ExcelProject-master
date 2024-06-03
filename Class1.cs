using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excel
{
    public static class Parser
    {
        public static void remove_spaces(ref string x)
        {
            x = String.Concat(x.Where(c => !Char.IsWhiteSpace(c)));
        }

        public static List<double> merge(dynamic x, dynamic y)
        {
            if (x is List<double> && y is List<double>)
            {
                x.AddRange(y);
                return x;
            }
            else if (x is List<double> && y is double)
            {
                x.Add(y);
                return x;
            }
            else if (x is double && y is List<double>)
            {
                y.Add(x);
                return y;
            }
            else if (x is double && y is double)
            {
                return new List<double> { x, y };
            }
            else
            {
                throw new ArgumentException("List");
            }
        }

        public static int find_bracket(string x, int i)

        {
            char y = x[i];

            if (y == '(')
            {
                y = ')';

                while (i < x.Length)
                {
                    if (x[i] == y)
                        return i;

                    ++i;
                }
            }

            if (y == ')')
            {
                y = '(';

                while (i >= 0)
                {
                    if (x[i] == y)
                        return i;

                    --i;
                }
            }

            if (y == '[')
            {
                y = ']';

                while (i < x.Length)
                {
                    if (x[i] == y)
                        return i;

                    ++i;
                }
            }

            if (y == ']')
            {
                y = '[';

                while (i >= 0)
                {
                    if (x[i] == y)
                        return i;

                    --i;
                }
            }

            return -1;
        }

        public static int find_left(string x, char y)
        {
            int i = 0;

            while (i < x.Length)
            {
                if (x[i] == y && i != 0 && x[i - 1] != '+' && x[i - 1] != '-' && x[i - 1] != '(' && x[i - 1] != '*' && x[i - 1] != '/')
                    return i;

                if (x[i] == ')' || x[i] == ']')
                    return -1;

                if (x[i] == '(' || x[i] == '[')
                {
                    i = find_bracket(x, i);

                    if (i == -1)
                        throw new ArgumentException("Brackets");
                }

                ++i;
            }

            return -1;
        }

        public static int find_right(string x, char y)
        {
            int i = x.Length - 1;

            while (i >= 0)
            {
                if (x[i] == y && i != 0 && x[i - 1] != '+' && x[i - 1] != '-' && x[i - 1] != '(' && x[i - 1] != '*' && x[i - 1] != '/')
                    return i;

                if (x[i] == '(' || x[i] == '[')
                    return -1;

                if (x[i] == ')' || x[i] == ']')
                {
                    i = find_bracket(x, i);

                    if (i == -1)
                        throw new ArgumentException("Brackets");
                }

                --i;
            }

            return -1;
        }

        public static dynamic parse(string x)
        {
            //remove_spaces(ref x);
            if (x.Contains("}") || x.Contains("{") || x.Contains("!"))
                throw new ArgumentException("Symbols");

            x = x.Replace(" ", "");
            x = x.Replace("mod", "%");
            x = x.Replace("div", "\\");
            x = x.Replace(">=", "}");
            x = x.Replace("<=", "{");
            x = x.Replace("<>", "!");
            x = x.Replace("and", "&");
            x = x.Replace("or", "|");
            x = x.Replace("eqv", "=");
            if (x[0] == '+')
                return parse(x.Substring(1, x.Length - 1));
            if (x[0] == '*' || x[0] == '/')
                throw new ArgumentException("Operators");

            if (x[0] == '(' && x[x.Length - 1] == ')')
                return parse(x.Substring(1, x.Length - 2));
            if (x[0] == '[' && x[x.Length - 1] == ']')
            {
                return parse(x.Substring(1, x.Length - 2));
            }

            Tuple<char, Func<dynamic, dynamic, dynamic>, bool>[] ops = new Tuple<char, Func<dynamic, dynamic, dynamic>, bool>[]
            {

                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('|', (dynamic a, dynamic b) => a || b, false),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('&', (dynamic a, dynamic b) => a && b, false),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('=', (dynamic a, dynamic b) => a == b, false),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('!', (dynamic a, dynamic b) => a != b, false),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('>', (dynamic a, dynamic b) => a > b, false),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('<', (dynamic a, dynamic b) => a < b, false),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('}', (dynamic a, dynamic b) => a >= b, false),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('{', (dynamic a, dynamic b) => a <= b, false),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('%', (dynamic a, dynamic b) => a - b * Math.Floor(a / b), true),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('\\', (dynamic a, dynamic b) => Math.Floor(a / b), true),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('+', (dynamic a, dynamic b) => a + b, false),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('-', (dynamic a, dynamic b) => a - b, true),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('*', (dynamic a, dynamic b) => a * b, false),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('/', (dynamic a, dynamic b) => a / b, true),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>('^', (dynamic a, dynamic b) => Math.Pow(a, b), true),
                new Tuple<char, Func<dynamic,dynamic,dynamic>, bool>(',', (dynamic a, dynamic b) => merge(a, b), false),
            };

            foreach (var i in ops)
            {
                int j;

                if (i.Item3)
                    j = find_right(x, i.Item1);
                else
                    j = find_left(x, i.Item1);

                if (j > 0)
                    return i.Item2(parse(x.Substring(0, j)), parse(x.Substring(j + 1, x.Length - j - 1)));
            }

            if (x[0] == '-')
            {
                return -parse(x.Substring(1, x.Length - 1));
            }
            else if (x.Length > 3 && x.Substring(0, 3) == "min")
            {
                List<double> f = parse(x.Substring(3, x.Length - 3));
                return f.Min();
            }
            else if (x.Length > 3 && x.Substring(0, 3) == "max")
            {
                List<double> f = parse(x.Substring(3, x.Length - 3));
                return f.Max();
            }
            else if (x.Length > 3 && x.Substring(0, 3) == "not")
            {
                return !parse(x.Substring(3, x.Length - 3));
            }
            else if (x == "true" || x == "false")
            {
                return x == "true";
            }
            else
            {
                return Convert.ToDouble(x, CultureInfo.InvariantCulture);
            }
        }
    };

    public class Cell 
    {
        public DataGridViewCell Example { get; set; }
        public dynamic Value { get; set; }
        public string Expression { get; set; }
        public bool Busy { get; set; }
        public bool Count { get; set;}

        public Cell(DataGridViewCell example, string exp)
        {
            this.Expression = exp;
            this.Example = example;
            Value = 0;
            this.Busy = false;
            this.Count = true;
        }

        public int process()
        {
            if (Busy)
            {
                throw new Exception("Recursion");
            }
            else
            {
                if (Count)
                {
                    return Value;
                }
                else
                {
                    Busy = true;
                    dynamic temp = Parser.parse(Expression);
                    Count = true;
                    Busy = false;
                    Value = temp;
                    return temp;
                }
            }
        }

    };
    public class CellCount
    {
        private static CellCount instance;
        public static CellCount Instance
        {
            get
            {
                if (instance == null) 
                {
                    instance = new CellCount();
                }
                return instance;
            }
        }
        private DataGridView table;

        public void Providetable(DataGridView table)
        {
            this.table = table;
        }
        public Cell TakeCell(int row, int column)
        {
            Cell cell = (Cell) table[column, row].Tag;
            return cell;
        }

        public Cell TakeCell(DataGridViewCell tablecell)
        {
            Cell cell = (Cell)tablecell.Tag;
            return cell;
        } 

        public bool existance(int rownum, int columnnum)
        {
            try 
            {
                TakeCell(rownum, columnnum);
                return true;
            }
            catch(ArgumentOutOfRangeException)
            {
                return false;
            }
        }

        public dynamic GetValue(int rownum, int columnnum)
        {
            if (existance(rownum, columnnum))
            {
                Cell cell = TakeCell(rownum, columnnum);
                if (cell.Expression != "")
                {
                    return cell.process();
                }
                else
                {
                    return cell.Value;
                }
            }
            else
            {
                throw new ArgumentOutOfRangeException("This cell does not exist");
            }
        }
    };

}
