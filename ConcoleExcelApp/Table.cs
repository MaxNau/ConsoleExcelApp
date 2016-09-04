using System;
using System.Data;
using System.Linq;

namespace ConcoleExcelApp
{
    public class Table
    {
        private readonly Cell[,] _table;

        public Table(int rows, int columns)
        {
            _table = new Cell[rows, columns];
        }

        // populates table with user input
        // if user doesn't fill all cells, each cell without input trited as empty
        public void PopulateTable()
        {
            // temporary variable for storing user input

            for (int col = 0; col < _table.GetLength(0); col++)
            {
                // read user input and store it in temporary variable
                var temp = Console.ReadLine();
                // split user input string into seprate strings that correspond to cells
                var nums = temp.Split(' ');

                    for (int row = 0; row < _table.GetLength(1); row++)
                    {
                        // check if user filled all the cells through input
                        // if check is valid then define cell data type
                        // else triet cell as empty
                        if (nums.Length > row)
                            _table[col, row] = CheckCellDataType(nums[row]);
                        else
                            _table[col, row] = CheckCellDataType("");

                        // define the index of the cell in Excel format
                        _table[col, row].Index = string.Format("{1}{0}", col + 1, NumberToAlpha(row));
                    }
                }
    
            DisplayTable();
            CalculateResult();
        }

        // determines the data type of the cell
        public Cell CheckCellDataType(string data)
        {
            Cell cell = new Cell(data);

            int iresult;

            // check if cell is empty
            if (data == "")
            {
                cell.DataType = CellDataType.Empty;
                return cell;
            }

            // check if cell contains digit, formula, text or cell contains invalid user input
            if (int.TryParse(data, out iresult))
            {
                cell.DataType = CellDataType.Numeric;
            }
            else if (data[0] == '=')
            {
                cell.DataType = CellDataType.Formula;
            }
            else if (data[0] == '\'')
            {
                cell.DataType = CellDataType.Text;
            }
            else
            {
                cell.DataType = CellDataType.Error;
                cell.Data = "#Invalid input";
                return cell;
            }

            return cell;
        }

        // transfors number to the english alphabet letter
        public string NumberToAlpha(long number, bool isLower = false)
        {
            string returnVal = "";
            char c = isLower ? 'a' : 'A';
            while (number >= 0)
            {
                returnVal = (char)(c + number % 26) + returnVal;
                number /= 26;
                number--;
            }

            return returnVal;
        }

        // searches for the links in cells and return the value that is located in corresponding cell
        public object SearchIndex(string index)
        {
            for (int c = 0; c < _table.GetLength(0); c++)
            {
                for (int r = 0; r < _table.GetLength(1); r++)
                {
                    // if index of the cell corresponds to the index that we search and cell data type is not a formula
                    // than return the data from the found cell
                    // else if index of the cell corresponds to the index that we search and cell data type is text
                    // then return the found cell
                    // else if index of the cell corresponds to the index that we search and cell data type is formula
                    // check if formula can be evaluated (all links are replaced with values)
                    // if formula can be evaluated return the result
                    // if not search for the links and replace them with values until formula can be evaluated

                    if (_table[c, r].Index == index)
                    {
                        if (_table[c, r].DataType == CellDataType.Numeric)
                        {
                            return _table[c, r].Data;
                        }
                        else if (_table[c, r].DataType == CellDataType.Text)
                        {
                            return _table[c, r].Data;
                        }
                        else if (_table[c, r].DataType == CellDataType.Formula)
                        {
                            if (CheckFormula(_table[c, r].Data))
                            {
                                _table[c, r].Data = Convert.ToDouble(new DataTable().Compute((_table[c, r].Data).Substring(1), null)).ToString();
                                _table[c, r].DataType = CellDataType.Numeric;
                                return _table[c, r].Data;
                            }
                            else
                            {
                                string[] indexes = _table[c, r].Data.Split('+', '-', '/', '*', '=');
                                foreach (string ind in indexes)
                                {
                                    if (ind != "")
                                    {
                                        if (char.IsLetter(ind[0]))
                                        {
                                            _table[c, r].Data = _table[c, r].Data.Replace(ind, (string)SearchIndex(ind));
                                            CheckCellDataType(_table[c, r].Data);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return SearchIndex(index);
        }


        // calculates the resulting cells
        public void CalculateResult()
        {
            for (int c = 0; c < _table.GetLength(0); c++)
            {
                for (int r = 0; r < _table.GetLength(1); r++)
                {
                    if (_table[c, r].DataType == CellDataType.Formula)
                    {
                        if (CheckFormula(_table[c, r].Data))
                        {
                            if (_table[c, r].DataType != CellDataType.Numeric)
                            {
                                if (_table[c, r].Data[0] == '=')
                                {
                                    _table[c, r].Data = Convert.ToDouble(new DataTable().Compute(_table[c, r].Data.Substring(1), null)).ToString();
                                    //table[c, r].dataType = CellDataType.numeric;
                                }
                                else
                                {
                                    _table[c, r].Data = Convert.ToDouble(new DataTable().Compute(_table[c, r].Data, null)).ToString();
                                    //table[c, r].dataType = CellDataType.numeric;
                                }
                            }
                            _table[c, r].DataType = CellDataType.Numeric;
                        }
                        else
                        {
                            string[] indexes = _table[c, r].Data.Split('+', '-', '/', '*', '=');
                            foreach(string ind in indexes)
                            {
                                if (ind != "")
                                {
                                    if (char.IsLetter(ind[0]))
                                    {
                                        _table[c, r].Data = _table[c, r].Data.Replace(ind, (string)SearchIndex(ind));
                                        if (_table[c, r].DataType == CellDataType.Text || _table[c, r].DataType == CellDataType.Numeric)
                                            _table[c, r].Data = _table[c, r].Data.Substring(1);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }


        // checks if cell doesn't contain any links to other cells and can be evaluated
        public bool CheckFormula(string formula)
        {
            return formula.All(c => Char.IsDigit(c) || c.Equals('+') || c.Equals('-') || 
            c.Equals('*') || c.Equals('/') || c.Equals('=') || c.Equals('.'));
        }

        // displayes the table
        public void DisplayTable(bool evaluated = false)
        {
            string outstring;
            Console.Write(Environment.NewLine);

            for (int c = 0; c < _table.GetLength(0); c++)
            {
                for (int r = 0; r < _table.GetLength(1); r++)
                {
                    if(_table[c, r].DataType == CellDataType.Text & evaluated)
                        outstring = (_table[c, r].Data).Substring(1);
                    else
                        outstring = _table[c, r].Data;
                    Console.Write($"{outstring,-15} ");
                }
                Console.Write(Environment.NewLine);
            }
        }
    }
}
