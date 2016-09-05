using System;
using System.Data;
using System.Globalization;
using System.Linq;

namespace ConcoleExcelApp
{
    public class Table
    {
        private readonly ExcelTable _excelTable;

        public Table(int rows, int columns)
        {
            _excelTable = new ExcelTable(rows, columns);
        }

        // inner class that holds the table
        private class ExcelTable
        {
            private const int ColumnIndocator = 0;
            private const int RowIndicator = 1;

            private readonly Cell[,] _table;

            public ExcelTable(int rows, int columns)
            {
                _table = new Cell[rows, columns];
            }

            public Cell this[int columnIndex, int rowIndex]
            {
                get { return _table[columnIndex, rowIndex]; }
                set { _table[columnIndex, rowIndex] = value; }
            }

            // Returns number of columns.
            public int GetColumnLength => _table.GetLength(ColumnIndocator);

            // Returns number of rows.
            public int GetRowLength => _table.GetLength(RowIndicator);
        }

        
        // Populates table with user input.
        // If user doesn't fill all cells - cells without input trited as empty.
        public void PopulateTable()
        {
            for (int col = 0; col < _excelTable.GetColumnLength; col++)
            {
                // Read user input and store it in temporary variable
                var userInput = Console.ReadLine();

                // Split user input string into seprate strings that correspond to row cells
                if (userInput != null)
                {
                    var cellData = userInput.Split(' ');

                    for (int row = 0; row < _excelTable.GetRowLength; row++)
                    {
                        // Check if user filled all the cells through input.
                        // If check is valid then define cell data type
                        // else triet cell as empty
                        if (cellData.Length > row)
                        {
                            _excelTable[col, row] = CheckCellDataType(cellData[row]);
                        }
                        else
                            _excelTable[col, row] = CheckCellDataType("");

                        // Define the index of the cell in Excel format
                        _excelTable[col, row].Index = string.Format("{1}{0}", col + 1, NumberToAlpha(row));
                    }
                }
            }
    
            DisplayTable();
            CalculateResult();
        }

        // Determines the data type of the cell
        private Cell CheckCellDataType(string data)
        {
            Cell cell = new Cell(data);

            int iresult;

            // Check if cell is empty
            if (string.IsNullOrEmpty(data))
            {
                cell.DataType = CellDataType.Empty;
                return cell;
            }

            // Check if cell contains digit, formula, text or cell contains invalid user input
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
                string error;

                if (!ValidateUserInput(cell, out error))
                {
                    cell.DataType = CellDataType.Error;
                    return cell;
                }
            }
            

            return cell;
        }

        // Transforms number to the english alphabet letter
        private string NumberToAlpha(long number, bool isLower = false)
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

        // Transforms english alphabet letter to number
        private int AlphaToNumber(string alpha)
        {
            int returnVal = 0;
            string col = alpha.ToUpper();
            for (int iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                char colPiece = col[iChar];
                int colNum = colPiece - 64;
                returnVal = returnVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
            }
            return returnVal;
        }

        // Searches for the links in cells and returns the value that is located in corresponding cell.
        // If index of the cell corresponds to the index that we search check the cell data type.
        // If cell data type is text or number then return the data in found cell.
        // Else, if cell data type is formula check if formula can be evaluated (all links are replaced with values).
        // If formula can be evaluated calculate and return cell data,
        // if not search for the links and replace them with values until formula can be evaluated.
        private object SearchIndex(string index)
        {
            for (int c = 0; c < _excelTable.GetColumnLength; c++)
            {
                for (int r = 0; r < _excelTable.GetRowLength; r++)
                {
                    if (_excelTable[c, r].Index == index)
                    {
                        if (_excelTable[c, r].DataType == CellDataType.Numeric || _excelTable[c, r].DataType == CellDataType.Text)
                        {
                            return _excelTable[c, r].Data;
                        }
                        else if (_excelTable[c, r].DataType == CellDataType.Formula)
                        {
                            if (CheckFormula(_excelTable[c, r].Data))
                            {
                                _excelTable[c, r].Data = Convert.ToDouble(new DataTable().Compute((_excelTable[c, r].Data).Substring(1), null)).ToString(CultureInfo.InvariantCulture);
                                _excelTable[c, r].DataType = CellDataType.Numeric;
                                return _excelTable[c, r].Data;
                            }
                            else
                            {
                                string[] indexes = _excelTable[c, r].Data.Split('+', '-', '/', '*', '=');
                                foreach (string ind in indexes)
                                {
                                    if (ind != "")
                                    {
                                        if (char.IsLetter(ind[0]))
                                        {
                                            _excelTable[c, r].Data = _excelTable[c, r].Data.Replace(ind, (string)SearchIndex(ind));
                                            CheckCellDataType(_excelTable[c, r].Data);
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

        // Calculates the resulting cells
        public void CalculateResult()
        {
            for (int c = 0; c < _excelTable.GetColumnLength; c++)
            {
                for (int r = 0; r < _excelTable.GetRowLength; r++)
                {
                    if (_excelTable[c, r].DataType == CellDataType.Formula)
                    {
                        if (CheckFormula(_excelTable[c, r].Data))
                        {
                            CalculateCell(_excelTable[c, r]);
                        }
                        else
                        {
                            FindAndReplaceFormulaLinks(_excelTable[c, r]);
                        }
                    }
                }
            }
        }

        // Calculates cell data
        private void CalculateCell(Cell cell)
        {
            if (cell.DataType != CellDataType.Numeric)
            {
                if (cell.Data[0] == '=')
                {
                    cell.Data = Convert.ToDouble(new DataTable().Compute(cell.Data.Substring(1), null)).ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    cell.Data = Convert.ToDouble(new DataTable().Compute(cell.Data, null)).ToString(CultureInfo.InvariantCulture);
                }
            }
            cell.DataType = CellDataType.Numeric;
        }

        // Replaces links in formula with values in corresponding linked cell
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


        // Checks if cell doesn't contain any links to other cells and can be evaluated
        private bool CheckFormula(string formula)
        {
            return formula.All(c => char.IsDigit(c) || c.Equals('+') || c.Equals('-') || 
            c.Equals('*') || c.Equals('/') || c.Equals('=') || c.Equals('.'));
        }

        // Displayes the table
        public void DisplayTable(bool evaluated = false)
        {
            Console.Write(Environment.NewLine);

            for (int c = 0; c < _excelTable.GetColumnLength; c++)
            {
                for (int r = 0; r < _excelTable.GetRowLength; r++)
                {
                    string outstring;

                    if(_excelTable[c, r].DataType == CellDataType.Text & evaluated)
                        outstring = (_excelTable[c, r].Data).Substring(1);
                    else
                        outstring = _excelTable[c, r].Data;
                    Console.Write($"{outstring,-15} ");
                }
                Console.Write(Environment.NewLine);
            }
        }

        // Validates user input
        private bool ValidateUserInput(Cell cell, out string error)
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

        // Validates formula
        private bool ValidateFormula(string [] formula, out string error)
        {
            error = "";

            foreach (var argument in formula)
            {
                if (!IsValideFormula(argument, out error))
                    return false;
            }

            return true;
        }

        // Validates formula arguments
        private bool IsValideFormula(string argument, out string error)
        {
            error = "";
            if (IsDigit(argument))
                return true;
            else if (ValidateIndex(argument, out error))
                return true;

            return false;
        }

        // Checks if data contains digit
        private bool IsDigit(string data)
        {
            int res;
            return int.TryParse(data, out res);
        }

        // Checks if link contains valid index
        private bool IsValidIndex(string index)
        {
            return char.IsLetter(index[0]);
        }

        // Validates link indexes and arguments
        private bool ValidateIndex(string index, out string error)
        {
            error = "";

            if (IsValidIndex(index))
            {
                string colIndex = new string(index.TakeWhile(char.IsLetter).ToArray());
                string rowIndex = new string(index.SkipWhile(char.IsLetter).ToArray());

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

        // Checks whether the link column index, specified by the user, exists within the table
        private bool CheckColumnIndex(int index)
        {
            return index <= _excelTable.GetColumnLength;
        }

        // Checks whether the link row index, specified by the user, exists within the table
        private bool CheckRowIndex(int index)
        {
            return index <= _excelTable.GetRowLength;
        }
    }
}
