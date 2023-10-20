import Excel, { Row } from 'exceljs';
import { isFunction, isEmpty } from '@/utils/is';
import { TableColumnAllProps } from '../../types';

type ReadExcelOptions = {
  sheetIndex?: number;
};

/* 将所有的行数据转换为json */
const changeRowsToDict = (
  worksheet: Excel.Worksheet,
  columns: Partial<TableColumnAllProps>[],
  headerRowNumber: number
) => {
  const dataArray = [];
  const validate = [];
  // 校验表头是否跟配置的一致
  const columnsTitle: string[] = columns.map(column => column.title as string);
  const headerRow: Row = worksheet.getRow(headerRowNumber);
  const rowValues = headerRow?.values;
  const length = rowValues.length as number;
  // 根据表头的title字典树
  const titleDict = getTitleDict(columns);

  // 校验导入的表格中是否属于表头数据
  for (let i = 1; i < length; i++) {
    if (!columnsTitle.includes(rowValues[i])) {
      return { data: dataArray, validate: ['导入的文件表头与模板不一致'] };
    }
  }

  // 遍历每一行， > headerRowNumber 判定为数据行
  worksheet.eachRow(function (row: Row, rowNumber: number) {
    if (rowNumber > headerRowNumber) {
      // 每一行的数据存储
      const data = {};

      row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        const columnTitle = rowValues[colNumber];
        const targetColumn = titleDict[columnTitle];
        // 获取每个单元格对应的column的配置
        const {
          required,
          title,
          validate: validateFunction,
          message,
          dataIndex,
          importHandler
        } = targetColumn || {};

        // 当前单元格为空
        if (isEmpty(cell.value as string)) {
          // 判断是否为必填字段
          if (required) {
            validate.push(`第${rowNumber}行：${title}必填`);
          }
        } else {
          // 单元格内容不为空且设置了 validate 函数，则需要对内容的有效性做检查
          if (isFunction(validateFunction)) {
            const errorMessage = `第${rowNumber}行：${title}${message}`;
            const result = validateFunction(cell.value);

            if (!result) {
              validate.push(errorMessage);
            }
          }
        }

        data[dataIndex] = isFunction(importHandler) ? importHandler(cell.value) : cell.value;
        data['rowNumber'] = rowNumber;
      });

      dataArray.push(data);
    }
  });

  return { data: dataArray, validate };
};

export const readExcel = (
  file,
  columns: Partial<TableColumnAllProps>[],
  options?: ReadExcelOptions
) => {
  // 获取所有column的叶子节点
  const leafNodes = [];
  // 获取最大深度，作为表头的最后一行
  const maxLevel = getMaxLevel(columns);

  getLeafNodes(columns, leafNodes);

  const sheetIndex = options?.sheetIndex || 1;

  return new Promise(async (resolve, reject) => {
    try {
      const workbook = new Excel.Workbook();
      const result = await workbook.xlsx.load(file);

      const worksheet = result.getWorksheet(sheetIndex);
      // 获取数据
      const dataArray = changeRowsToDict(worksheet, leafNodes, maxLevel);

      resolve(dataArray);
    } catch (e) {
      reject(e);
    }
  });
};

// 获取所有叶子节点
function getLeafNodes(
  headerColumn: Partial<TableColumnAllProps>[],
  leafNodes: Partial<TableColumnAllProps>[]
) {
  for (let i = 0; i < headerColumn.length; i++) {
    const item = headerColumn[i];

    if (item.children) {
      getLeafNodes(item.children, leafNodes);
    } else {
      leafNodes.push(item);
    }
  }
}

// 获取columns层级深度
function getMaxLevel(arr: any): number {
  let maxLevel = 1;

  function traverse(arr: any, level: number) {
    for (let i = 0; i < arr.length; i++) {
      const obj = arr[i];

      if (obj.children) {
        traverse(obj.children, level + 1);
      }
    }
    maxLevel = Math.max(maxLevel, level);
  }

  traverse(arr, 1);
  return maxLevel;
}

/**
 * 根据title获取对应的column
 * @param columns
 * @returns
 */
function getTitleDict(columns: Partial<TableColumnAllProps>[]) {
  const ans = {};

  for (let i = 0; i < columns.length; i++) {
    const column = columns[i];

    ans[column.title as string] = column;
  }
  return ans;
}
