using ClosedXML.Excel;
using System;
using System.IO;


namespace ExcelMacroRun
{
    /// <summary>
    /// Excelの読み込みとワークシート、セルへのアクセスを行うクラス
    /// </summary
    class ExcelLoader : IDisposable
    {
        #region メンバ変数
        /// <summary>
        /// Excelのワークブック(2013以降のExcelを指定すること)
        /// </summary>
        private XLWorkbook _Workbook = null;
        #endregion

        #region プロパティ
        /// <summary>
        /// 全てのワークシートを取得します。
        /// </summary>
        public IXLWorksheets Worksheets
        {
            get
            {
                if (this._Workbook != null)
                {
                    return this._Workbook.Worksheets;
                }
                else
                {
                    return null;
                }
            }
        }
        #endregion

        #region コンストラクタ
        /// <summary>
        /// 新しいクラスのインスタンスを作成します。
        /// </summary>
        public ExcelLoader()
        {
        }

        /// <summary>
        /// 新しいクラスのインスタンスを作成します。
        /// </summary>
        /// <param name="path">Excelファイルのパス</param>
        public ExcelLoader(string path)
        {
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                // ファイルパスが設定されているかつ、そのパスのファイルが存在する場合、ファイルを開く
                this._Workbook = new XLWorkbook(path);
            }
        }
        #endregion

        #region 公開メソッド
        /// <summary>
        /// 解放処理
        /// </summary>
        public void Dispose()
        {
            CloseExcel();
        }

        /// <summary>
        /// 引数で指定されたExcelファイルを開きます。
        /// </summary>
        /// <param name="path">Excelファイルのパス</param>
        /// <returns>true：オープン成功、false：オープン失敗</returns>
        public bool OpenExcel(string path)
        {
            CloseExcel();

            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                // ファイルパスが設定されているかつ、そのパスのファイルが存在する場合、ファイルを開く
                this._Workbook = new XLWorkbook(path);
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 指定したインデックスのワークシートの指定セルを読み込みます。
        /// </summary>
        /// <param name="row">行番号</param>
        /// <param name="col">列番号</param>
        /// <param name="index">ワークシートのインデックス番号</param>
        /// <returns>true：オープン成功、false：オープン失敗</returns>
        public string ReadCell(int row, int col, int index = 1)
        {
            if (this._Workbook == null)
            {
                return string.Empty;
            }

            try
            {
                var worksheet = this._Workbook.Worksheet(index);
                var cell = worksheet.Cell(row, col);
                return cell.Value as String;
            }
            catch
            {
                // 指定されたワークシートが存在しなかった場合、空値を返却する
                return string.Empty;
            }
        }

        /// <summary>
        /// 指定したワークシートの指定セルを読み込みます。
        /// </summary>
        /// <param name="row">行番号</param>
        /// <param name="col">列番号</param>
        /// <param name="worksheet">ワークシートのインデックス番号</param>
        /// <returns>true：オープン成功、false：オープン失敗</returns>
        public string ReadCell(int row, int col, IXLWorksheet worksheet)
        {
            if (worksheet == null)
            {
                return string.Empty;
            }

            try
            {
                var cell = worksheet.Cell(row, col);
                return cell.Value as String;
            }
            catch
            {
                // 指定されたワークシートのセルが存在しなかった場合、空値を返却する
                return string.Empty;
            }
        }

        /// <summary>
        /// 現在オープンされているExcelファイルをクローズします。
        /// </summary>
        public void CloseExcel()
        {
            if (this._Workbook != null)
            {
                // 既にExcelが開かれている場合、破棄する
                this._Workbook.Dispose();
                this._Workbook = null;
            }
        }
        #endregion
    }
}
