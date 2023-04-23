import ExcelJS, { Column, Cell } from "exceljs";
import { ICellStyle, IExportToXlsxConfig } from "./IExportToXlsx";



const ApplyCellStyle = (cell: Cell, config: ICellStyle) => {
  const { borderProps, bgColor, fontProps, ...rest } = config;
  //fill prop section
  if (bgColor) {
    cell.fill = {
      ...cell.fill,
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: bgColor },
    };
  }
  if (rest.fill) cell.fill = rest.fill;
  //fill prop section />

  //font prop section
  if (fontProps)
    cell.font = {
      ...cell.font,
      color: fontProps.color ? { argb: fontProps.color } : undefined,
      size: fontProps.size,
      bold: fontProps.bold,
      ...rest.font,
    };
  //fill prop section />

  //border prop section
  if (borderProps)
    cell.border = {
      bottom: {
        color: borderProps.color ? { argb: borderProps.color } : undefined,
        style: borderProps.variant,
      },
      left: {
        color: borderProps.color ? { argb: borderProps.color } : undefined,
        style: borderProps.variant,
      },
      right: {
        color: borderProps.color ? { argb: borderProps.color } : undefined,
        style: borderProps.variant,
      },
      top: {
        color: borderProps.color ? { argb: borderProps.color } : undefined,
        style: borderProps.variant,
      },
      ...rest.border,
    };
    //fill prop section />
};

/**
 * @param {string} key keys from the data object
 */
export interface IColumnDef<T = any> extends Partial<Omit<Column, "key">> {
  key: keyof T;
}

export const ExportToXlsx = <T = { [key: string]: string | number }>(
  data: T[],
  columnsDef: IColumnDef<T>[],
  config?: IExportToXlsxConfig
) => {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet("my sheet");

  const { setCellStyle, setHeaderStyle, filename, cellStyles, headerStyles } = config || {};

  const hasSomeHeaders = columnsDef?.find((x) => x.header !== undefined)
    ? true
    : false;

  sheet.columns = columnsDef as Partial<Column>[];

  // write the data
  for (const item of data) sheet.addRow(item);

  //style for the rest of the rows
  if (setCellStyle != undefined || cellStyles != undefined) {
    for (let i = 1; i <= sheet.rowCount; i++) {
      sheet.getRow(i).eachCell((c, i) => {
        if (setCellStyle) setCellStyle(c);
        if (cellStyles) ApplyCellStyle(c, cellStyles);
      });
    }
  }

  //style for header
  if (
    hasSomeHeaders &&
    (setHeaderStyle !== undefined || headerStyles != undefined)
  ) {
    sheet.getRow(1).eachCell((c, i) => {
      if (setHeaderStyle) setHeaderStyle(c);
      if (headerStyles) ApplyCellStyle(c, headerStyles);
    });
  }

  wb.xlsx.writeBuffer().then((data) => {
    const blob = new Blob([data], {
      type: "application/vnd.openxmlformats-officedocument.spreedsheet.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename
      ? filename.replace(".xlsx", "") + ".xlsx"
      : "download.xlsx";
    anchor.click();
    window.URL.revokeObjectURL(url);
  });
};
