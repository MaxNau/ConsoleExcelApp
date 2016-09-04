using System;
using System.Data;
using System.Linq;
using Microsoft.SqlServer.Server;

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
                    {
                        _table[col, row] = CheckCellDataType(nums[row]);
                    }
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

            string error;

            if (!ValidateUserInput(cell, out error))
            {
                cell.DataType = CellDataType.Error;
                cell.Data = error;
                return cell;
            }

            return cell;
        }

        // transfors alpha to the english alphabet letter
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

        public int AlphaToNumber(string alpha, bool isLower = false)
        {
            int retVal = 0;
            string col = alpha.ToUpper();
            for (int iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                char colPiece = col[iChar];
                int colNum = colPiece - 64;
                retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
            }
            return retVal;
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
                        if (_table[c, r].DataType == CellDataType.Numeric || _table[c, r].DataType == CellDataType.Text)
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
                            CalculateCell(_table[c, r]);
                            /*if (_table[c, r].DataType != CellDataType.Numeric)
                            {
                                if (_table[c, r].Data[0] == '=')
                                {
                                    _table[c, r].Data = Convert.ToDouble(new DataTable().Compute(_table[c, r].Data.Substring(1), null)).ToString();
                                }
                                else
                                {
                                    _table[c, r].Data = Convert.ToDouble(new DataTable().Compute(_table[c, r].Data, null)).ToString();
                                }
                            }
                            _table[c, r].DataType = CellDataType.Numeric;*/
                        }
                        else
                        {
                            FindAndReplaceFormulaLinks(_table[c, r]);
                            /*string[] indexes = _table[c, r].Data.Split('+', '-', '/', '*', '=');
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
                            }*/
                        }
                    }
                }
            }
        }

        private void CalculateCell(Cell cell)
        {
            if (cell.DataType != CellDataType.Numeric)
            {
                if (cell.Data[0] == '=')
                {
                    cell.Data = Convert.ToDouble(new DataTable().Compute(cell.Data.Substring(1), null)).ToString();
                }
                else
                {
                    cell.Data = Convert.ToDouble(new DataTable().Compute(cell.Data, null)).ToString();
                }
            }
            cell.DataType = CellDataType.Numeric;
        }

        private void FindAndReplaceFormulaLinks(Cell cell)
        {
            string[] indexes = cell.Data.Split('+', '-', '/', '*', '=');
            foreach (string index in indexes)
            {
                if (index != "")
                {
                    if (char.IsLetter(index[0]))
                    {
                        cell.Data = cell.Data.Replace(index, (string)SearchIndex(index));
                        if (cell.DataType == CellDataType.Text || cell.DataType == CellDataType.Numeric)
                            cell.Data = cell.Data.Substring(1);
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

        public bool ValidateUserInput(Cell cell, out string error)
        {
            error = "";

            if (cell.DataType == CellDataType.Formula)
            {
                return ValidateFormula(cell.Data.Substring(1).Split('+', '-', '/', '*'), out error);
            }
            else if (cell.DataType == CellDataType.Numeric)
            {
                return IsDigit(cell.Data);
            }

            return false;
        }

        public bool ValidateFormula(string [] formula, out string error)
        {
            error = "";

            foreach (var argument in formula)
            {
                if (!IsValideFormula(argument, out error))
                    return false;
            }

            return true;
        }

        public bool IsValideFormula(string argument, out string error)
        {
            error = "";
            if (IsDigit(argument))
                return true;
            else if (ValidateIndex(argument, out error))
                return true;

            return false;
        }

        public bool IsDigit(string data)
        {
            int res;
            return int.TryParse(data, out res);
        }

        public bool IsValidIndex(string index)
        {
            return char.IsLetter(index[0]);
        }

        public bool ValidateIndex(string index, out string error)
        {
            error = "";

            if (IsValidIndex(index))
            {
                string colIndex = new String(index.TakeWhile(Char.IsLetter).ToArray());
                string rowIndex = new String(index.SkipWhile(Char.IsLetter).ToArray());

                if (rowIndex != "")
                {
                    if (index.Length == colIndex.Length + rowIndex.Length)
                    {
                        if (CheckColumnIndex(AlphaToNumber(colIndex)) & CheckRowIndex(int.Parse(rowIndex)))
                            return true;
                    }
                }
            }

            error = "#Invalid cell link in formula";
            return false;
        }

        public bool CheckColumnIndex(int index)
        {
            return index <= _table.GetLength(0);
        }

        public bool CheckRowIndex(int index)
        {
            return index <= _table.GetLength(1);
        }
    }
}
