# GitDocumentos
Lluvia de ideas para rutinas generales

--Mensajes de advertencia.
--Mensajes de error.
--Mensajes de confirmacion.
--Mensajes .
-- Excel

public class ManejoExcell
{
    private int xlRight = -4152;
    private int xlLeft = -4131;
    private int xlBottom = 4107;
    private int xlTop = -4160;
    private int xlCenter = -4108;
    private int xlUnderlineStyleSingle = 2;
    private string _formatofecha;
    public string formatofecha
    {
        get
        {
            return _formatofecha;
        }

        set
        {
            _formatofecha = value;
        }
    }

    public object AbrirExcel(bool xvisible = false)
    {
        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
        object OleApp;
        OleApp = CreateObject("Excel.Application");
        OleApp.Visible = xvisible;
        OleApp.Workbooks.Add();
        return OleApp;
    }

    public void AbreHoja(ref object OleApp)
    {
        OleApp.Sheets.Add();
    }

    public void Muestra(ref object OleApp)
    {
        OleApp.visible = true;
    }

    public void Esconde(ref object OleApp)
    {
        OleApp.visible = false;
    }

    public void GuardaArchivoExcel(ref object Oleapp, string Directorioynombre)
    {
        Oleapp.ActiveWorkbook.SaveAs(Directorioynombre);
        Oleapp.Quit();
    }

    public void AutoFit(object OleApp)
    {
        OleApp.Cells.Select();
        OleApp.Cells.EntireColumn.AutoFit();
        OleApp.Range("A1:A1").Select();
    }

    public string CrearRangoExcel(int Fila, int Columna, int CantidadColumnas = 0)
    {
        string Rango;
        if (CantidadColumnas > 0)
            Rango = Trim(ValorLetras(Columna, 0)) + Trim(Str(Fila)) + ":" + Trim(ValorLetras(CantidadColumnas, Columna)) + Trim(Str(Fila));
        else
            Rango = Trim(ValorLetras(Columna, 0)) + Trim(Str(Fila));
        return Rango;
    }

    public string ValorLetras(int CantidadColumnas, int ColumnaInicial)
    {
        string Caracter;
        if (CantidadColumnas > 26)
        {
            Caracter = Chr(Int(CantidadColumnas / 26) + Asc("A") - 1 + ColumnaInicial) + Chr((CantidadColumnas - (Int(CantidadColumnas / 26) * 26)) + Asc("A") - 1 + ColumnaInicial);
            if (CantidadColumnas % 26 == 0)
            {
                Caracter = Caracter.Replace("@", "Z");
                Caracter = Caracter.Replace(Caracter.Substring(0, 1), Chr(Asc(Caracter.Substring(0, 1)) - 1));
            }
        }
        else
            Caracter = Chr(CantidadColumnas + Asc("A") - 1 + ColumnaInicial);
        return Caracter;
    }

    public void Titulo(ref object OleApp, string Titulo, int Fila, int Columna, int CantidadColumnas = 0, string TipoLetra = "Arial", int Tamano = 9, bool Negrita = false, bool Subrayado = false, bool Italica = false)
    {
        OleApp.Cells(Fila, Columna).Value = Titulo;
        OleApp.Range(CrearRangoExcel(Fila, Columna, CantidadColumnas - 1)).Select();
        if (Subrayado)
            OleApp.Selection.Font.Underline = xlUnderlineStyleSingle;
        OleApp.Selection.Font.Name = TipoLetra;
        OleApp.Selection.Font.Size = Tamano;
        OleApp.Selection.Font.Italic = Italica;
        OleApp.Selection.Font.Bold = Negrita;
        OleApp.Selection.HorizontalAlignment = xlCenter;
        OleApp.Selection.VerticalAlignment = xlTop;
        OleApp.Selection.WrapText = true;
        OleApp.Selection.Orientation = 0;
        OleApp.Selection.ShrinkToFit = false;
        OleApp.Selection.MergeCells = true;
    }

    public void Celda(ref object OleApp, object Titulo, int Fila, int Columna)
    {
        OleApp.Cells(Fila, Columna).Value = Titulo;
    }

    public void CeldaFormato(ref object OleApp, string Formato, int Fila1, int Columna, int Fila2, string TipoLetra = "Arial", int Tamano = 9, bool Negrita = false, bool Subrayado = false, bool Italica = false, bool ParaCantidad = false, bool Fmt_cant_inv = false, bool MonedaDolar = false)
    {
        OleApp.Range(CrearRangoExcel(Fila1, Columna, 0) + ":" + CrearRangoExcel(Fila2, Columna, 0)).Select();
        string pc_maskdec_inv = ".99";
        string pc_maskdecimales = "";
        switch (Trim(UCase(Formato)))
        {
            case "NUMERO":
            {
                OleApp.Selection.HorizontalAlignment = xlRight;
                OleApp.selection.NumberFormat = "#.##0,00";
                break;
            }

            case "CARACTER":
            {
                OleApp.Selection.NumberFormat = "@";
                OleApp.Selection.HorizontalAlignment = xlLeft;
                break;
            }

            case "ENTERO":
            {
                OleApp.Selection.NumberFormat = "0;[Red]-0";
                OleApp.Selection.HorizontalAlignment = xlRight;
                break;
            }

            case "ENTERO2":
            {
                OleApp.Selection.NumberFormat = "0;[Red]-0";
                OleApp.selection.NumberFormat = "#.##0";
                break;
            }

            case "FECHA":
            {
                OleApp.Selection.NumberFormat = "dd/mm/yy;@";
                OleApp.Selection.HorizontalAlignment = xlRight;
                break;
            }

            case "FECHA1":
            {
                OleApp.Selection.NumberFormat = "dd-mm-yy;@";
                OleApp.Selection.HorizontalAlignment = xlRight;
                break;
            }

            case "FECHA2":
            {
                OleApp.Selection.NumberFormat = "dd-mm-yy HH:mm;@";
                OleApp.Selection.HorizontalAlignment = xlRight;
                break;
            }

            case "FECHA3":
            {
                OleApp.Selection.NumberFormat = "dd/MM/yyyy;@";
                OleApp.Selection.HorizontalAlignment = xlRight;
                break;
            }

            case "FECHA4":
            {
                OleApp.Selection.NumberFormat = "dd/MM/yyyy HH:mm;@";
                OleApp.Selection.HorizontalAlignment = xlRight;
                break;
            }

            case "FECHA5":
            {
                OleApp.Selection.NumberFormat = "MM/dd/yyyy;@";
                OleApp.Selection.HorizontalAlignment = xlRight;
                break;
            }

            case "FECHA6":
            {
                OleApp.Selection.NumberFormat = "MM/dd/yyyy HH:mm;@";
                OleApp.Selection.HorizontalAlignment = xlRight;
                break;
            }

            case "FECHA7":
            {
                OleApp.Selection.NumberFormat = "MM/dd/yy;@";
                OleApp.Selection.HorizontalAlignment = xlRight;
                break;
            }

            case "FECHA8":
            {
                OleApp.Selection.NumberFormat = "MM/dd/yy HH:mm;@";
                OleApp.Selection.HorizontalAlignment = xlRight;
                break;
            }
        }

        if (Subrayado)
            OleApp.Selection.Font.Underline = xlUnderlineStyleSingle;
        OleApp.Selection.Font.Name = TipoLetra;
        OleApp.Selection.Font.Size = Tamano;
        OleApp.Selection.Font.Italic = Italica;
        OleApp.Selection.Font.Bold = Negrita;
        OleApp.Selection.VerticalAlignment = xlTop;
        OleApp.Selection.WrapText = false;
        OleApp.Selection.Orientation = 0;
        OleApp.Selection.ShrinkToFit = false;
    }

    public void CeldaFormatoDouble(ref object OleApp, string Formato, int Fila1, int Columna, int Fila2, string TipoLetra = "Arial", int Tamano = 9, bool Negrita = false, bool Subrayado = false, bool Italica = false, bool ParaCantidad = false, bool Fmt_cant_inv = false, bool MonedaDolar = false)
    {
        OleApp.Range(CrearRangoExcel(Fila1, Columna, 0) + ":" + CrearRangoExcel(Fila2, Columna, 0)).Select();
        OleApp.Selection.NumberFormat = "###,###,###,##0.00_);[Red](###,###,###,##0.00)";
        if (Subrayado)
            OleApp.Selection.Font.Underline = xlUnderlineStyleSingle;
        OleApp.Selection.Font.Name = TipoLetra;
        OleApp.Selection.Font.Size = Tamano;
        OleApp.Selection.Font.Italic = Italica;
        OleApp.Selection.Font.Bold = Negrita;
        OleApp.Selection.VerticalAlignment = xlTop;
        OleApp.Selection.WrapText = false;
        OleApp.Selection.Orientation = 0;
        OleApp.Selection.ShrinkToFit = false;
    }

    public void SumaColumna(ref object OleApp, int FilaResultado, int ColumnaResultado, int Columna, int FilaDesde, int FilaHasta, bool MonedaDolar = false, bool sumcashocant = false, bool Fmt_cant_inv = false)
    {
        OleApp.Range(CrearRangoExcel(FilaResultado, ColumnaResultado, 0)).Select();
        OleApp.ActiveCell.FormulaR1C1 = CrearFormula(FilaResultado, ColumnaResultado, Columna, FilaDesde, FilaHasta);
        string pc_maskdec_inv = "";
        string pc_maskdecimales = "";
        if (sumcashocant)
            if (Fmt_cant_inv)
            {
                if (Trim(pc_maskdec_inv) == "")
                    OleApp.Selection.NumberFormat = "#,###,##0_);[Red](#,##0.00)";
                if (Trim(pc_maskdec_inv) == ".9")
                    OleApp.Selection.NumberFormat = "#,###,##0.0_);[Red](#,##0.00)";
                if (Trim(pc_maskdec_inv) == ".99")
                    OleApp.Selection.NumberFormat = "#,###,##0.00_);[Red](#,##0.00)";
                if (Trim(pc_maskdec_inv) == ".999")
                    OleApp.Selection.NumberFormat = "#,###,##0.000_);[Red](#,##0.00)";
                if (Trim(pc_maskdec_inv) == ".9999")
                    OleApp.Selection.NumberFormat = "#,###,##0.0000_);[Red](#,##0.00)";
            }
            else
            {
                if (Trim(pc_maskdecimales) == "")
                    OleApp.Selection.NumberFormat = "#,###,##0_);[Red](#,##0.00)";
                if (Trim(pc_maskdecimales) == ".9")
                    OleApp.Selection.NumberFormat = "#,###,##0.0_);[Red](#,##0.00)";
                if (Trim(pc_maskdecimales) == ".99")
                    OleApp.Selection.NumberFormat = "#,###,##0.00_);[Red](#,##0.00)";
                if (Trim(pc_maskdecimales) == ".999")
                    OleApp.Selection.NumberFormat = "#,###,##0.000_);[Red](#,##0.00)";
                if (Trim(pc_maskdecimales) == ".9999")
                    OleApp.Selection.NumberFormat = "#,###,##0.0000_);[Red](#,##0.00)";
            }
        else if (MonedaDolar)
            OleApp.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)";
        else
            OleApp.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)";
        OleApp.Selection.Font.Bold = true;
    }

    public string CrearFormula(int FilaResultado, int ColumnaResultado, int Columna, int FilaDesde, int FilaHasta)
    {
        string resultado;
        resultado = "=SUM(R[" + Trim(Str(FilaDesde - FilaResultado)) + "]C[" + Trim(Str(Columna - ColumnaResultado)) + "]" + ":R[" + Trim(Str(FilaHasta - FilaResultado)) + "]C[" + Trim(Str(Columna - ColumnaResultado)) + "])";
        return resultado;
    }

    public void Indentar(ref object OleApp, int Fila, int Columna, int Cantidad)
    {
    }

    public void FExpExcel(DataSet ds, string ptitulo, string tabla)
    {
        string ctitulo, pc_titulo1, pc_titulo2, pc_titulo3;
        int row, col;
        object ole_excel;
        System.TypeCode tipo;
        object valor;
        pc_titulo1 = "";
        pc_titulo2 = "";
        pc_titulo3 = "";
        if (ds.Tables(tabla).Rows.Count > 0)
        {
            ole_excel = AbrirExcel(true);
            ctitulo = Trim(ptitulo);
            Titulo(ref ole_excel, ctitulo, 2, 2, 9, "Bookman Old Style", 20, false, false, false);
            if ((pc_titulo1 != null))
                Titulo(ref ole_excel, Trim(pc_titulo1), 2, 1, 9, "Arial Back", 15, true, false, false);
            if ((pc_titulo2 != null))
                Titulo(ref ole_excel, Trim(pc_titulo2), 3, 1, 9, "Arial Back", 15, true, false, false);
            if ((pc_titulo3 != null))
                Titulo(ref ole_excel, Trim(pc_titulo3), 4, 1, 9, "Arial Back", 15, true, false, false);
            for (col = 0; col <= ds.Tables(tabla).Columns.Count - 1; col++)
            {
                CeldaCaracter(ref ole_excel, UCase(Trim(ds.Tables(tabla).Columns(col).ColumnName)), 5, col + 1, "Arial", 9, true, false, false);
                for (row = 0; row <= ds.Tables(tabla).Rows.Count - 1; row++)
                {
                    valor = ds.Tables(tabla).Rows(row).Item(col);
                    tipo = valor.GetTypeCode;
                    Celda(ref ole_excel, valor, row + 6, col + 1);
                }

                if (tipo == System.TypeCode.String)
                    CeldaFormato(ref ole_excel, "CARACTER", 5, col + 1, ds.Tables(tabla).Rows.Count + 5);
                if (tipo == System.TypeCode.DateTime)
                    CeldaFormato(ref ole_excel, "FECHA", 5, col + 1, ds.Tables(tabla).Rows.Count + 5);
                if (tipo == System.TypeCode.Int16 | tipo == System.TypeCode.Int32 | tipo == System.TypeCode.Int64)
                    CeldaFormato(ref ole_excel, "ENTERO", 5, col + 1, ds.Tables(tabla).Rows.Count + 5);
                if (tipo == System.TypeCode.Double | tipo == System.TypeCode.Decimal)
                    CeldaFormato(ref ole_excel, "NUMERO", 5, col + 1, ds.Tables(tabla).Rows.Count + 5);
            }

            AutoFit(ole_excel);
            Muestra(ref ole_excel);
        }
    }

    public void CeldaCaracter(ref object OleApp, string pTitulo, int Fila, int Columna, string TipoLetra = "Arial", int Tamano = 9, bool Negrita = false, bool Subrayado = false, bool Italica = false)
    {
        OleApp.Range(CrearRangoExcel(Fila, Columna, 0)).Select();
        OleApp.Selection.NumberFormat = "@";
        OleApp.Cells(Fila, Columna).Value = Trim(pTitulo);
        if (Subrayado)
            OleApp.Selection.Font.Underline = xlUnderlineStyleSingle;
        OleApp.Selection.Font.Name = TipoLetra;
        OleApp.Selection.Font.Size = Tamano;
        OleApp.Selection.Font.Italic = Italica;
        OleApp.Selection.Font.Bold = Negrita;
        OleApp.Selection.HorizontalAlignment = xlLeft;
        OleApp.Selection.VerticalAlignment = xlTop;
        OleApp.Selection.WrapText = false;
        OleApp.Selection.Orientation = 0;
        OleApp.Selection.ShrinkToFit = false;
    }

    public void CeldaNumero(ref object OleApp, string pTitulo, int Fila, int Columna, string TipoLetra = "Arial", int Tamano = 9, bool Negrita = false, bool Subrayado = false, bool Italica = false)
    {
        OleApp.Cells(Fila, Columna).Value = Trim(pTitulo);
        OleApp.Range(CrearRangoExcel(Fila, Columna, 0)).Select();
        OleApp.Selection.NumberFormat = "#,##0.00";
        OleApp.Selection.style = "comma";
        if (Subrayado)
            OleApp.Selection.Font.Underline = xlUnderlineStyleSingle;
        OleApp.Selection.Font.Name = TipoLetra;
        OleApp.Selection.Font.Size = Tamano;
        OleApp.Selection.Font.Italic = Italica;
        OleApp.Selection.Font.Bold = Negrita;
        OleApp.Selection.HorizontalAlignment = xlRight;
        OleApp.Selection.VerticalAlignment = xlTop;
        OleApp.Selection.WrapText = false;
        OleApp.Selection.Orientation = 0;
        OleApp.Selection.ShrinkToFit = false;
    }

    public void CeldaEntero(ref object OleApp, string pTitulo, int Fila, int Columna, string TipoLetra = "Arial", int Tamano = 9, bool Negrita = false, bool Subrayado = false, bool Italica = false)
    {
        OleApp.Cells(Fila, Columna).Value = Trim(pTitulo);
        OleApp.Range(CrearRangoExcel(Fila, Columna, 0)).Select();
        OleApp.Selection.NumberFormat = "#,##0";
        if (Subrayado)
            OleApp.Selection.Font.Underline = xlUnderlineStyleSingle;
        OleApp.Selection.Font.Name = TipoLetra;
        OleApp.Selection.Font.Size = Tamano;
        OleApp.Selection.Font.Italic = Italica;
        OleApp.Selection.Font.Bold = Negrita;
        OleApp.Selection.HorizontalAlignment = xlRight;
        OleApp.Selection.VerticalAlignment = xlTop;
        OleApp.Selection.WrapText = false;
        OleApp.Selection.Orientation = 0;
        OleApp.Selection.ShrinkToFit = false;
    }

    public void CeldaFecha(ref object OleApp, string Titulo, int Fila, int Columna, string TipoLetra = "Arial", int Tamano = 9, bool Negrita = false, bool Subrayado = false, bool Italica = false)
    {
        OleApp.Cells(Fila, Columna).Value = Titulo;
        OleApp.Range(CrearRangoExcel(Fila, Columna, 0)).Select();
        OleApp.Selection.NumberFormat = "dd-mm-yy;@";
        if (Subrayado)
            OleApp.Selection.Font.Underline = xlUnderlineStyleSingle;
        OleApp.Selection.Font.Name = TipoLetra;
        OleApp.Selection.Font.Size = Tamano;
        OleApp.Selection.Font.Italic = Italica;
        OleApp.Selection.Font.Bold = Negrita;
        OleApp.Selection.HorizontalAlignment = xlRight;
        OleApp.Selection.VerticalAlignment = xlTop;
        OleApp.Selection.WrapText = false;
        OleApp.Selection.Orientation = 0;
        OleApp.Selection.ShrinkToFit = false;
    }
}

