import { Cell, BorderStyle } from "exceljs";


/**
 * @param cellStyles global styles
 * additionalProps: fontColor, borderColor, borderVariant, bgColor
 * @param headerStyles global styles
 * additionalProps: fontColor, borderColor, borderVariant, bgColor
 */
export interface IExportToXlsxConfig {
    filename?: string;
    cellStyles?: ICellStyle;
    headerStyles?: ICellStyle;
    setCellStyle?: (cell: Cell) => void;
    setHeaderStyle?: (cell: Cell) => void;
  }
  
  export interface ICellStyle extends Partial<Cell> {
    fontProps?: IFont;
    borderProps?: IBorder;
    bgColor?: string;
  }
  
  export interface IFont {
    color?: string;
    size?: number;
    bold?: boolean;
  }
  
  export interface IBorder {
    color?: string;
    variant?: BorderStyle;
  }