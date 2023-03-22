using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics.CodeAnalysis;
using Excel = Microsoft.Office.Interop.Excel;


namespace MatrixMultiplication
{
    class Program
    {
        static void Main(string[] args)
        {
            int[,] matrixA = new int[,] { { 1, 2, 3 }, { 4, 5, 6 }, { 7, 8, 9 } };
            int[,] matrixB = new int[,] { { 4, 2, 1 }, { 7, 3, 5 }, { 6, 2, 3 } };

            int[,] resultPeremnoj = new int[3, 3];


            for (int i = 0; i < matrixA.GetLength(0); i++)
            {
                for (int j = 0; j < matrixB.GetLength(1); j++)
                {
                    int res = 0;

                    for (int k = 0; k < matrixA.GetLength(1); k++)
                    {
                        res += matrixA[i, k] * matrixB[k, j];
                    }

                    resultPeremnoj[i, j] = res;

                }
            }

            // Создайте экземпляр приложения Excel
            Excel.Application excelApp = new Excel.Application();

            // Создает экземпляр книги Excel и открывает его из предопределенного местоположения
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();

            // Добавляет новый лист в книгу с именем Datatable.
            Excel.Worksheet excelWorkSheet = (Excel.Worksheet)excelWorkBook.ActiveSheet;
            excelWorkSheet.Name = "Matrix";

            // Цвет границ А
            Excel.Range cellA = excelWorkSheet.Range["B2", "D4"];
            cellA.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cellA.Borders.ColorIndex = 4;
            cellA.Borders.Weight = Excel.XlBorderWeight.xlThick;

            // Цвет границ В
            Excel.Range cellB = excelWorkSheet.Range["F2", "H4"];
            cellB.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cellB.Borders.ColorIndex = 4;
            cellB.Borders.Weight = Excel.XlBorderWeight.xlThick;

            // Цвет границ результата
            Excel.Range cellRes = excelWorkSheet.Range["J2", "L4"];
            cellRes.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cellRes.Borders.ColorIndex = 3;
            cellRes.Borders.Weight = Excel.XlBorderWeight.xlThick;

            // Имена столбцов
            excelWorkSheet.Cells[1, 2] = "Матрица A";
            excelWorkSheet.Cells[1, 6] = "Матрица B";
            excelWorkSheet.Cells[1, 10] = "Результат матриц А * В, методом перемножения";
            excelWorkSheet.Cells[6, 10] = "Результат среднего арифметического значения";
            excelWorkSheet.Cells[10, 10] = "Результат дисперсии матрицы";
            excelWorkSheet.Cells[14, 10] = "Результат мат ожидания матрицы";

            // Имена строк
            excelWorkSheet.Cells[2, 1] = "Строка 1";
            excelWorkSheet.Cells[3, 1] = "Строка 2";
            excelWorkSheet.Cells[4, 1] = "Строка 3";
            excelWorkSheet.Cells[7, 1] = "Строка 4";
            excelWorkSheet.Cells[8, 1] = "Строка 5";
            excelWorkSheet.Cells[11, 1] = "Строка 6";
            excelWorkSheet.Cells[12, 1] = "Строка 7";
            excelWorkSheet.Cells[15, 1] = "Строка 8";
            excelWorkSheet.Cells[16, 1] = "Строка 9";


            // A
            excelWorkSheet.Cells[2, 2] = matrixA[0, 0];
            excelWorkSheet.Cells[2, 3] = matrixA[0, 1];
            excelWorkSheet.Cells[2, 4] = matrixA[0, 2];
            excelWorkSheet.Cells[3, 2] = matrixA[1, 0];
            excelWorkSheet.Cells[3, 3] = matrixA[1, 1];
            excelWorkSheet.Cells[3, 4] = matrixA[1, 2];
            excelWorkSheet.Cells[4, 2] = matrixA[2, 0];
            excelWorkSheet.Cells[4, 3] = matrixA[2, 1];
            excelWorkSheet.Cells[4, 4] = matrixA[2, 2];

            // B
            excelWorkSheet.Cells[2, 6] = matrixB[0, 0];
            excelWorkSheet.Cells[2, 7] = matrixB[0, 1];
            excelWorkSheet.Cells[2, 8] = matrixB[0, 2];
            excelWorkSheet.Cells[3, 6] = matrixB[1, 0];
            excelWorkSheet.Cells[3, 7] = matrixB[1, 1];
            excelWorkSheet.Cells[3, 8] = matrixB[1, 2];
            excelWorkSheet.Cells[4, 6] = matrixB[2, 0];
            excelWorkSheet.Cells[4, 7] = matrixB[2, 1];
            excelWorkSheet.Cells[4, 8] = matrixB[2, 2];

            // result peremnoj
            excelWorkSheet.Cells[2, 10] = resultPeremnoj[0, 0];
            excelWorkSheet.Cells[2, 11] = resultPeremnoj[0, 1];
            excelWorkSheet.Cells[2, 12] = resultPeremnoj[0, 2];
            excelWorkSheet.Cells[3, 10] = resultPeremnoj[1, 0];
            excelWorkSheet.Cells[3, 11] = resultPeremnoj[1, 1];
            excelWorkSheet.Cells[3, 12] = resultPeremnoj[1, 2];
            excelWorkSheet.Cells[4, 10] = resultPeremnoj[2, 0];
            excelWorkSheet.Cells[4, 11] = resultPeremnoj[2, 1];
            excelWorkSheet.Cells[4, 12] = resultPeremnoj[2, 2];

            // Расчет среднего по строкам
            Excel.Range srdRes1 = excelWorkSheet.Range["J2", "L2"];
            double srdResMatric1 = excelWorkSheet.Application.WorksheetFunction.Average(srdRes1);
            excelWorkSheet.Cells[7, 10] = srdResMatric1;

            Excel.Range srdRes2 = excelWorkSheet.Range["J3", "L3"];
            double srdResMatric2 = excelWorkSheet.Application.WorksheetFunction.Average(srdRes2);
            excelWorkSheet.Cells[7, 11] = srdResMatric2;

            Excel.Range srdRes3 = excelWorkSheet.Range["J4", "L4"];
            double srdResMatric3 = excelWorkSheet.Application.WorksheetFunction.Average(srdRes3);
            excelWorkSheet.Cells[7, 12] = srdResMatric3;

            // Расчет среднего по столбцам
            Excel.Range srdRes11 = excelWorkSheet.Range["J2", "J4"];
            double srdResMatric11 = excelWorkSheet.Application.WorksheetFunction.Average(srdRes11);
            excelWorkSheet.Cells[8, 10] = srdResMatric11;

            Excel.Range srdRes22 = excelWorkSheet.Range["K2", "K4"];
            double srdResMatric22 = excelWorkSheet.Application.WorksheetFunction.Average(srdRes22);
            excelWorkSheet.Cells[8, 11] = srdResMatric22;

            Excel.Range srdRes33 = excelWorkSheet.Range["L2", "L4"];
            double srdResMatric33 = excelWorkSheet.Application.WorksheetFunction.Average(srdRes33);
            excelWorkSheet.Cells[8, 12] = srdResMatric33;

            // Цвет границ результата среднего арифметического 
            Excel.Range sr1 = excelWorkSheet.Range["J7"];
            sr1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            sr1.Borders.ColorIndex = 5;
            sr1.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range sr2 = excelWorkSheet.Range["K7"];
            sr2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            sr2.Borders.ColorIndex = 5;
            sr2.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range sr3 = excelWorkSheet.Range["L7"];
            sr3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            sr3.Borders.ColorIndex = 5;
            sr3.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range sr11 = excelWorkSheet.Range["J8"];
            sr11.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            sr11.Borders.ColorIndex = 5;
            sr11.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range sr22 = excelWorkSheet.Range["K8"];
            sr22.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            sr22.Borders.ColorIndex = 5;
            sr22.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range sr33 = excelWorkSheet.Range["L8"];
            sr33.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            sr33.Borders.ColorIndex = 5;
            sr33.Borders.Weight = Excel.XlBorderWeight.xlThick;

            // Дисперсия по строкам и столбцам
            Excel.Range d1 = excelWorkSheet.Range["J2", "L2"];
            Excel.Range d2 = excelWorkSheet.Range["J3", "L3"];
            Excel.Range d3 = excelWorkSheet.Range["J4", "L4"];
            Excel.Range d11 = excelWorkSheet.Range["J2", "J4"];
            Excel.Range d22 = excelWorkSheet.Range["K2", "K4"];
            Excel.Range d33 = excelWorkSheet.Range["L2", "L4"];

            double disp1 = excelApp.WorksheetFunction.Var_S(d1);
            double disp2 = excelApp.WorksheetFunction.Var_S(d2);
            double disp3 = excelApp.WorksheetFunction.Var_S(d3);
            double disp11 = excelApp.WorksheetFunction.Var_S(d11);
            double disp22 = excelApp.WorksheetFunction.Var_S(d22);
            double disp33 = excelApp.WorksheetFunction.Var_S(d33);

            excelWorkSheet.Cells[11, 10] = disp1;
            excelWorkSheet.Cells[11, 11] = disp2;
            excelWorkSheet.Cells[11, 12] = disp3;
            excelWorkSheet.Cells[12, 10] = disp11;
            excelWorkSheet.Cells[12, 11] = disp22;
            excelWorkSheet.Cells[12, 12] = disp33;

            // Дисперсия - дизайн
            Excel.Range dp1 = excelWorkSheet.Range["J11"];
            dp1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            dp1.Borders.ColorIndex = 6;
            dp1.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range dp2 = excelWorkSheet.Range["K11"];
            dp2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            dp2.Borders.ColorIndex = 6;
            dp2.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range dp3 = excelWorkSheet.Range["L11"];
            dp3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            dp3.Borders.ColorIndex = 6;
            dp3.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range dp11 = excelWorkSheet.Range["J12"];
            dp11.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            dp11.Borders.ColorIndex = 6;
            dp11.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range dp22 = excelWorkSheet.Range["K12"];
            dp22.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            dp22.Borders.ColorIndex = 6;
            dp22.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range dp33 = excelWorkSheet.Range["L12"];
            dp33.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            dp33.Borders.ColorIndex = 6;
            dp33.Borders.Weight = Excel.XlBorderWeight.xlThick;

            // Мат ожидание по строкам и столбцам
            Excel.Range m1 = excelWorkSheet.Range["J2", "L2"];
            Excel.Range m2 = excelWorkSheet.Range["J3", "L3"];
            Excel.Range m3 = excelWorkSheet.Range["J4", "L4"];
            Excel.Range m11 = excelWorkSheet.Range["J2", "J4"];
            Excel.Range m22 = excelWorkSheet.Range["K2", "K4"];
            Excel.Range m33 = excelWorkSheet.Range["L2", "L4"];

            double matWait1 = excelApp.WorksheetFunction.SumProduct(m1);
            double matWait2 = excelApp.WorksheetFunction.SumProduct(m2);
            double matWait3 = excelApp.WorksheetFunction.SumProduct(m3);
            double matWait11 = excelApp.WorksheetFunction.SumProduct(m11);
            double matWait22 = excelApp.WorksheetFunction.SumProduct(m22);
            double matWait33 = excelApp.WorksheetFunction.SumProduct(m33);

            excelWorkSheet.Cells[15, 10] = matWait1;
            excelWorkSheet.Cells[15, 11] = matWait2;
            excelWorkSheet.Cells[15, 12] = matWait3;
            excelWorkSheet.Cells[16, 10] = matWait11;
            excelWorkSheet.Cells[16, 11] = matWait22;
            excelWorkSheet.Cells[16, 12] = matWait33;

            // Мат ожидание - дизайн
            Excel.Range mt1 = excelWorkSheet.Range["J15"];
            mt1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            mt1.Borders.ColorIndex = 7;
            mt1.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range mt2 = excelWorkSheet.Range["K15"];
            mt2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            mt2.Borders.ColorIndex = 7;
            mt2.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range mt3 = excelWorkSheet.Range["L15"];
            mt3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            mt3.Borders.ColorIndex = 7;
            mt3.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range mt11 = excelWorkSheet.Range["J16"];
            mt11.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            mt11.Borders.ColorIndex = 7;
            mt11.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range mt22 = excelWorkSheet.Range["K16"];
            mt22.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            mt22.Borders.ColorIndex = 7;
            mt22.Borders.Weight = Excel.XlBorderWeight.xlThick;

            Excel.Range mt33 = excelWorkSheet.Range["L16"];
            mt33.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            mt33.Borders.ColorIndex = 7;
            mt33.Borders.Weight = Excel.XlBorderWeight.xlThick;

            // Save and close
            excelWorkBook.SaveAs(@"C:\VisualStudio\Vs_Projects\C#\PR_Titov\PR_TITOV_\PR_TITOV_\pr2\pr2_matrica");
            excelWorkBook.Close();
            excelApp.Quit();
            Console.WriteLine("Excel файл создан");
            Console.ReadKey();
        }
    }
}
