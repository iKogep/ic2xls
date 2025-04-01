using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO; //Файлы и папки
using System.Diagnostics; //Запуск программ
using Word = Microsoft.Office.Interop.Word; //Работа с Word
using Excel = Microsoft.Office.Interop.Excel; //Работа с Excel
using RegExp = System.Text.RegularExpressions; //Работа с регулярными выражениями

namespace ic2xls
{
    public partial class Form1 : Form
    {
        #region Глобальные переменные
        string _globalTXT = "";
        List<string> _tempList = new List<string> { }; //Целевые папки.
        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        #region Кнопка - открыть файл Word
        private void button_Open_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "MS Word documents|*.do*";
            fileDialog.Title = "Выбрать документ Word";
            if (fileDialog.ShowDialog() != DialogResult.OK) return;
            textBox_FileName.Text = fileDialog.FileName;
            //button_Convert.Enabled = true;
            button_Save.Enabled = true;
        }
        #endregion

        #region Кнопка - открыть файл Excel
        private void button_Save_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "MS Excel workbooks|*.xl*";
            fileDialog.Title = "Выбрать таблицу Excel";
            if (fileDialog.ShowDialog() != DialogResult.OK) return;
            textBox_FileExport.Text = fileDialog.FileName;
            button_Convert.Enabled = true;
            numericUpDown_Sheet.Enabled = true;
        }
        #endregion

        #region Кнопка - преобразовать
        private void button_Convert_Click(object sender, EventArgs e)
        {
            #region Открываем файл с данными RO
            try { File.Open(textBox_FileName.Text, FileMode.Open, FileAccess.Read, FileShare.None).Close(); }
            catch { MessageBox.Show("Выбранный файл недоступен.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            Microsoft.Office.Interop.Word.Application objWord = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document objDoc = new Microsoft.Office.Interop.Word.Document();
            objDoc = objWord.Documents.Open(textBox_FileName.Text);
            #endregion

            #region Получаем текст
            _globalTXT = objDoc.Range().Text;
            #endregion

            #region Закрываем файл
            objDoc.Close();
            objWord.Quit();
            #endregion

            #region Подготавливаем таблицу и систематизируем сведения о её формате
            _tempList.AddRange(_globalTXT.Split('\r')); //Создаем список - обрабатываемый текст построчно
            //подчищаем файл с хвоста
            const string _eot = "+------+------+-------+-----+--------+-------+------+";
            while (!_tempList[_tempList.Count - 1].StartsWith(_eot)) _tempList.RemoveAt(_tempList.Count - 1); //ищем лищний текст после метки конца таблицы и удаляем его
            if (_tempList[_tempList.Count - 1].StartsWith(_eot)) _tempList.RemoveAt(_tempList.Count - 1); //ищем метку конца таблицы и удаляем её (эту строку)
            //подчищаем файл с головы
            while (!_tempList[0].StartsWith("+-")) _tempList.RemoveAt(0); //удаляем лишний текст перед таблицей
            //определяем диапазон заголовка (графошапки), предполагая что в его структуре не может быть ошибки
            int _start = 0; //начало заголовка
            int _stop = 0; //конец заголовка
            for (int i = 1; i < _tempList.Count; i++) if (_tempList[i].StartsWith("+-")) _stop = i; //for (int i = 1; i < _tempList.Count; i++) if (_tempList[i].StartsWith("+-")) { _stop = i; i = _tempList.Count; }
            if (_stop == 0)
            {
                MessageBox.Show("Не удалось обнаружить заголовок (графошапку) таблицы.\r\nПроверьте корректность данных.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
                //можно добавить поиск обычного разделителя
            }
            //определяем размеры столбцов и создаем стандартный разделитель
            string _separator = ":-"; //разделитель строк
            string _tmpStr = _tempList[_stop].Replace('+', ':'); //в заголовке (графошапке) попа с разделителями, поэтому приводим к общему знаменателю
            List<string> _tmpStr0 = new List<string> { };
            _tmpStr0.AddRange(_tmpStr.Split(':'));
            int _columnCount = _tmpStr0.Count - 2; //количество столбцов в таблице
            int[,] _columnSize = new int[_columnCount, 2]; //свойства столбцов таблицы: 0 - начальная позиция; 1 - количество знаков.
            _columnSize[0, 0] = 1; //так как в позиции 0 находится разделитель //размер первого столбца
            _columnSize[0, 1] = _tmpStr0[1].Length; //количество знаков = длине строки
            for (int j = 2; j <= _tmpStr0.Count - 2; j++) //размеры столбцов, начиная со второго
            {
                _columnSize[j - 1, 0] = _columnSize[j - 2, 0] + _columnSize[j - 2, 1] + 1;
                _columnSize[j - 1, 1] = _tmpStr0[j].Length;
            }
            _separator = _tmpStr.Substring(0, _columnSize[_columnCount - 1, 0] + _columnSize[_columnCount - 1, 1] + 1); //настоящий разделитель
            //исправляем ошибки структуры, возможные при выгрузке данных из источника
            for (int i = 2; i < _tempList.Count - 1; i++) if (_tempList[i].Length > 0 && _tempList[i][0] != ':' && _tempList[i][0] != '+' && _tempList[i + 1].Length > 0 && _tempList[i + 1][0] == ':' && !(_tempList[i - 1].StartsWith(":-") || _tempList[i - 1].StartsWith("+-"))) _tempList[i] = _separator + _tempList[i]; //ищем строки с ошибками в таличной разметке и приводим их к формату нащего фильтра
            //подчищаем незначащие текстовые строки внутри таблицы
            for (int i = 0; i < _tempList.Count; ) if (_tempList[i].Length == 0 || !(_tempList[i][0] == '+' || _tempList[i][0] == ':')) _tempList.RemoveAt(i); else i++; //удаляем незначащие (не соответствующие табличному формату) строки
            #endregion

            #region Open XLSM
            try { File.Open(textBox_FileExport.Text, FileMode.Open, FileAccess.Write, FileShare.None).Close(); }
            catch { MessageBox.Show("Файл для экспорта недоступен для записи.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkBook = ObjExcel.Workbooks.Open(textBox_FileExport.Text);
            ObjExcel.Calculation = Excel.XlCalculation.xlCalculationManual; //Отключаем пересчет формул.
            #endregion

            ObjWorkSheet = ObjWorkBook.Sheets[numericUpDown_Sheet]; //выбираем номер листа

            #region Переносим заголовок - пустышка
            //заголовок не переносим, сложившийся факт!
            #endregion

            #region Переносим содержимое таблицы
            int _curRow = 4; //текущая строка на листе эксель в которую выгружаем данные, начинаем с 4 строки, 1-3 оставляем под оформление исполнителю
            bool _flag = true;
            _start = _stop + 1; //задаем начальную строку для экспорта
            string[] _expStr = new string[_columnCount]; //эту строку будем выгружать в эксель
            while (_flag)
            {
                for (int i = 0; i < _columnCount; i++) _expStr[i] = _tempList[_start].Substring(_columnSize[i, 0], _columnSize[i, 1]);
                for (int i = _start + 1; i < _tempList.Count; i++)
                {
                    if (!_tempList[i].StartsWith(_separator)) for (int j = 0; j < _columnCount; j++) _expStr[j] += _tempList[i].Substring(_columnSize[j, 0], _columnSize[j, 1]);
                    else
                    {
                        _stop = i;
                        _start = _stop + 1;
                        i = _tempList.Count;
                        if (checkBox_DateConvert.CheckState == CheckState.Checked) for (int j = 0; j < _columnCount; j++) ObjWorkSheet.Cells[_curRow, j + 1] = ICDateConverter(_expStr[j].TrimEnd());
                        else for (int j = 0; j < _columnCount; j++) ObjWorkSheet.Cells[_curRow, j + 1] = _expStr[j].TrimEnd();
                        _curRow++;
                    }
                }
                if (_start >= _tempList.Count || _stop >= _tempList.Count) _flag = false;
            }
            #endregion

            #region Save XLSM
            ObjExcel.Calculation = Excel.XlCalculation.xlCalculationAutomatic; //Включаем пересчет формул.
            ObjWorkBook.Save();
            ObjWorkBook.Close();
            ObjExcel.Quit();
            #endregion

            _tempList.Clear();

            #region Открываем готовый файл (папку с файлом)
            //try { System.Diagnostics.Process.Start(textBox_FileExport.Text); }
            try { System.Diagnostics.Process.Start("" + Directory.GetParent(textBox_FileName.Text)); }
            catch { MessageBox.Show("Файл недоступен.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            #endregion
        }
        #endregion

        #region Кнопка - выход
        private void button_Exit_Click(object sender, EventArgs e)
        {
            Close();
        }
        #endregion

        #region Автоматическое преобразование полей с датами даты
        private string ICDateConverter(string _input)
        {
            if (_input.Length != 10) return _input;
            if (_input == "   0  0  0") return "";
            if (RegExp.Regex.IsMatch(_input, "^\\d{4} \\d{2} \\d{2}")) return _input.Substring(8, 2) + "." + _input.Substring(5, 2) + "." + _input.Substring(0, 4);
            if (RegExp.Regex.IsMatch(_input, "^\\d{4} \\d{2}  \\d{1}")) return "0" + _input.Substring(9, 1) + "." + _input.Substring(5, 2) + "." + _input.Substring(0, 4);
            if (RegExp.Regex.IsMatch(_input, "^\\d{4}  \\d{1} \\d{2}")) return _input.Substring(8, 2) + ".0" + _input.Substring(6, 1) + "." + _input.Substring(0, 4);
            if (RegExp.Regex.IsMatch(_input, "^\\d{4}  \\d{1}  \\d{1}")) return "0" + _input.Substring(9, 1) + ".0" + _input.Substring(6, 1) + "." + _input.Substring(0, 4);
            return _input;
        }
        #endregion
    }
}
/* Регулярные выражения
Модель	Описание
^	                Начните с начала строки.
\s*	                Соответствует нулю или нескольким символам пробела.
[\+-]?	            Совпадение с нулевым или одним вхождением знака плюс или минус.
\s?	                Совпадение с нулем или одним символом пробела.
\$?	                Совпадение с нулевым или одним вхождением знака доллара.
\s?	                Совпадение с нулем или одним символом пробела.
\d*	                Соответствует нулю или нескольким десятичным числам.
\.?	                Совпадение с нулем или одним символом десятичной запятой.
\d{2}?	            Совпадение двух десятичных цифр ноль или один раз.
(\d*\.?\d{2}?){1}	Соответствие шаблону целой и дробной цифр, разделенных символом десятичной запятой, по крайней мере один раз.
$	                Совпадение с концом строки.
*/