/**
 * Excel-Deno-Bridge Type Definitions
 * Версия: 1.1.0
 */

/** Прокси-объект, представляющий любой объект Excel (Application, Range, Workbook и т.д.) */
interface ExcelProxy {
  /** Доступ к любому свойству или методу Excel */
  [key: string]: any;

  /** Позволяет вызывать объект как функцию (например, Excel.Range("A1")) */
  (...args: any[]): ExcelProxy & any;

  /** Свойство-маркер, указывающее, что это объект Excel */
  readonly __isExcelObject: boolean;

  // --- Общие свойства Excel для подсказок ---
  Value: any;
  Text: string;
  Formula: string;
  Address: string;
  Row: number;
  Column: number;

  /** Управление интерьером (фоном) ячейки */
  Interior: {
    /** Цвет в формате 0xBBGGRR */
    Color: number;
    ColorIndex: number;
    Pattern: number;
  };

  /** Управление шрифтом ячейки */
  Font: {
    Color: number;
    Bold: boolean;
    Italic: boolean;
    Size: number;
    Name: string;
  };

  /** Смещение относительно текущей ячейки */
  Offset(rowOffset: number, columnOffset: number): ExcelProxy;

  /** Выбор ячеек по адресу или индексу */
  Range(address: string): ExcelProxy;
  Cells(row: number, column: number): ExcelProxy;
}

/** Основной объект Excel (Application) */
declare const Excel: ExcelProxy & {
  ActiveCell: ExcelProxy;
  ActiveSheet: ExcelProxy;
  ActiveWorkbook: ExcelProxy;
  Selection: ExcelProxy;
  StatusBar: string | boolean;
  UserName: string;
  Version: string;
  DisplayAlerts: boolean;
  ScreenUpdating: boolean;
};

/** Глобальные функции моста */
declare global {
  /** Версия моста */
  const bridgeVersion: string;

  /**
   * Загружает и исполняет внешний JS файл из папки с книгой.
   * @param filename Имя файла (например, 'utils.js')
   */
  function include(filename: string): any;

  /**
   * Запускает фоновый мониторинг событий Excel в Rust.
   */
  function startExcelEvents(): "ok";

  /**
   * Синхронный сетевой запрос (блокирует поток).
   */
  function fetchSync(url: string): string;

  /**
   * Асинхронный сетевой запрос (рекомендуется).
   */
  function fetch(url: string): Promise<{
    text(): Promise<string>;
    json(): Promise<any>;
  }>;

  /**
   * Обработчик событий из Excel. Назначьте свою функцию.
   * @param name Имя события (например, 'cell_change')
   * @param data Данные события (значение ячейки)
   */
  var onExcelEvent: (name: string, data: string) => void;

  /**
   * Внутренняя функция опроса очереди событий.
   */
  function __runPoll(): { name: string; data: string } | null;
}

export {};
