// import { message } from 'antd';
import XLSX, { Worksheet } from 'exceljs';
import { isFunction } from '../is';
import saveAs from 'file-saver';
import { message } from 'antd';
import { TableColumnAllProps } from '../../types';

/**
 * excel 导出
 * style:excel表的样式配置
 * tableData:表的数据内容
 * headerColumns:表头配置
 * sheetName：工作表名
 */

export interface ExcelParamsType {
  sheetsName?: string[];
  headerStyle?: Partial<XLSX.Style>;
  style?: Partial<XLSX.Style>;
  headerColumns: TableColumnAllProps[][];
  tableDatas: any[][];
  fileName: string;
  isWorker?: boolean;
  ColumnWidth?: number;
}

const DEFAULT_COLUMN_WIDTH = 40;

export const exportExcel = async ({
  sheetsName,
  // 默认样式
  style = {
    alignment: {
      horizontal: 'center',
      vertical: 'middle'
    },
    font: {
      size: 14
    }
  },
  headerColumns,
  tableDatas,
  fileName,
  headerStyle,
  isWorker = true,
  ColumnWidth = DEFAULT_COLUMN_WIDTH
}: ExcelParamsType) => {
  if (!fileName) {
    const fileNameErrContent = '请传入文件名称';

    isWorker ? postMessage('error', fileNameErrContent) : message.error(fileNameErrContent);
    return;
  }
  // 如果没有传入sheet名称，默认采用fileName
  if (!sheetsName) {
    const fileNamePrefix = fileName.includes('.') ? fileName.split('.')[0] : fileName;

    sheetsName = [fileNamePrefix];
  }
  // sheetsName、columns、tableData数量是否对的上
  const isSheetCountRight =
    sheetsName.length === headerColumns.length && headerColumns.length === tableDatas.length;

  // sheet长度跟表格、表头长度对不上
  if (!isSheetCountRight) {
    const errorCntContent = 'sheet数量跟columns、tableData数量不匹配';

    isWorker ? postMessage('error', errorCntContent) : message.error(errorCntContent);
    return;
  }

  // 创建工作簿
  const workbook = new XLSX.Workbook();

  // 设置创建者
  workbook.creator = '我隔这敲代码呢';
  // 设置创建时间
  workbook.created = new Date();

  // 遍历sheet长度，依次给每个sheet添加表头，数据
  sheetsName.forEach((sheetName, index) => {
    // 添加工作表
    const worksheet = workbook.addWorksheet(sheetName);
    // 当前工作表对应的表头
    const headerColumn = headerColumns[index];
    // 当前工作表对应的数据
    const tableData = tableDatas[index];
    // 叶子节点，用于最后数据的匹配
    const leafNodes = [];

    // 给每个child设置parent, 同时获取叶子节点
    setChildParent(headerColumn, null, leafNodes);
    // children层级，以及每个层级对应的字段
    const { totalDepth, mergedColumns } = getMergedColumns(headerColumn, ColumnWidth);

    // 处理表头
    handleHeader(mergedColumns, worksheet, totalDepth, headerStyle);

    // 处理数据
    // 最后用于添加的数据
    const dataLength = handleData(tableData, leafNodes, worksheet);

    // 获取每列数据，依次对齐
    worksheet.columns.forEach(column => {
      column.alignment = style?.alignment as Partial<XLSX.Alignment>;
      column.font = style?.font;
      column.width = ColumnWidth;
    });
    // 设置数据的边框
    const tableRows = worksheet.getRows(totalDepth, dataLength + totalDepth);

    tableRows.forEach(row => {
      row.eachCell({ includeEmpty: true }, cell => {
        cell.border = style.border as Partial<XLSX.Borders>;
      });
    });
  });

  workbook.xlsx.writeBuffer().then(buffer => {
    const blobData = new Blob([buffer], { type: '' });

    isWorker ? postMessage('success', blobData) : saveAs(blobData, fileName);
  });
};

function handleData(tableData: any[], leafNodes: TableColumnAllProps[], worksheet: Worksheet) {
  const data = [];

  // 处理每个单元格的数据
  for (let i = 0; i < tableData.length; i++) {
    const dataRow = tableData[i];
    const pushedData = [];

    leafNodes.forEach((column, index) => {
      // 处理render情况
      const originValue = dataRow?.[column.dataIndex as string];
      let result = originValue;

      if (isFunction(column.render)) {
        // 有render则使用render返回结果，如果不是数字或者字符串，则使用原始数据展示
        result = column.render(dataRow?.[column.dataIndex as string], dataRow, index);
        const isStringOrNumber = typeof result === 'string' || typeof result === 'number';

        // 如果不是数字或者字符串，处理为原始数据
        if (!isStringOrNumber) {
          result = originValue;
        }
      }
      pushedData.push(result);
    });
    data.push(pushedData);
  }

  // 添加行
  if (data) worksheet.addRows(data);
  return data.length;
}

/**
 * 广度优先遍历，获取所需信息
 * @param columns
 * @returns
 */
function getMergedColumns(columns: TableColumnAllProps[], ColumnWidth: number) {
  let queue = [...columns];
  let tmpQueue = [];
  let depth = 0;
  const mergedColumns: TableColumnAllProps[][] = [];

  while (queue.length) {
    const column = queue.shift();

    if (!mergedColumns?.[depth]) {
      mergedColumns[depth] = [];
    }
    // 获取当前层级的节点
    const curLevelColumns = mergedColumns?.[depth];
    // 获取当前层级的前一个节点
    const lastColumn = curLevelColumns?.[curLevelColumns.length - 1];
    // 根据前一个节点判断当前节点的startIndex
    const startIndexByLastColumn = lastColumn?.endIndex ? lastColumn?.endIndex + 1 : 1;
    // 若当前节点为第一层的节点，则startIndex为startIndexByLastColumn，
    // 否则根据当前节点的父节点跟lastColumn的父节点是否是同一个节点判断当前节点的startIndex应该取父节点的startIndex还是startIndexByLastColumn
    const curStartIndex = !column.parent
      ? startIndexByLastColumn
      : column?.parent?.dataIndex === lastColumn?.parent?.dataIndex
        ? startIndexByLastColumn
        : column.parent.startIndex;

    column.startIndex = curStartIndex;

    // 获取当前节点的叶子节点的长度，用于填充二维数组
    const childrenLength = getChildLength(column.children, 0);

    const pushedColumn: TableColumnAllProps = {
      ...column,
      width: Math.floor((column.width as number) / 5) || ColumnWidth,
      startIndex: curStartIndex,
      endIndex: curStartIndex + (childrenLength ? childrenLength - 1 : 0),
      // 当前节点的深度，用于合并单元格
      depth: depth + 1,
      // 当前节点的叶子节点的长度，用于填充二维数组
      childrenLength: getChildLength(column.children, 0)
    };

    mergedColumns?.[depth]?.push(pushedColumn);

    if (column.children) {
      tmpQueue.push(...column.children);
    }

    if (queue.length === 0) {
      // 当前层节点遍历完成，切换到下一层
      if (tmpQueue.length) {
        queue = tmpQueue;
        tmpQueue = [];
        depth++;
      }
    }
  }

  return {
    totalDepth: depth + 1,
    mergedColumns
  };
}

function handleHeader(
  mergedColumns: TableColumnAllProps[][],
  worksheet: Worksheet,
  totalDepth: number,
  headerStyle: Partial<XLSX.Style>
) {
  // 获取总长度，即第一行节点的childLength 之和
  const totalLength = mergedColumns?.[0]?.reduce((total, item) => total + item.childrenLength, 0);
  // 构建一个二位数组，用 ‘’ 占位
  const rows = new Array(totalDepth).fill(0).map(() => new Array(totalLength).fill(''));

  // 找到每个节点对应的下标，进行数据表头数据的填充
  mergedColumns.forEach((columns, index) => {
    columns.forEach(item => {
      rows[index][item?.startIndex - 1] = item.title;
    });
  });
  // 添加表头
  const headerRows = worksheet.addRows(rows);

  headerRows.forEach(row => {
    row.eachCell((cell, column) => {
      // 设置背景色
      // cell.fill = {
      //   type: 'pattern',
      //   pattern: 'solid',
      //   fgColor: { argb: 'dff8ff' }
      // };
      // 设置字体
      cell.font = {
        bold: true,
        italic: false,
        size: 14,
        name: '微软雅黑',
        color: { argb: '000' }
      };
      cell.style = { ...cell.style, ...headerStyle };
    });
  });

  // 合并单元格
  mergeColumns(mergedColumns, worksheet, totalDepth);
}

/**
 * 合并单元格
 * @param columns 表头数据
 * @param worksheet worksheet实例
 * @param totalDepth 总深度
 */
function mergeColumns(columns: TableColumnAllProps[][], worksheet: Worksheet, totalDepth: number) {
  // 遍历每个column
  columns.forEach(column => {
    column.forEach(item => {
      // 数字转换为类似AA AB的字母，配合depth获取起始单元格的坐标
      const startCell = `${convertToTitle(item.startIndex)}${item.depth}`;
      let endCell = `${convertToTitle(item.endIndex)}${item.depth}`;

      // startIndex !== endIndex 说明当前单元格需要列合并
      if (item.startIndex !== item.endIndex) {
        worksheet.mergeCells(startCell + ':' + endCell);
      }
      // 如果当前单元格没有子节点，并且当前单元格所在层级不是最后一层，说明当前单元格需要行合并
      if (!item?.children?.length && item.depth !== totalDepth) {
        endCell = `${convertToTitle(item.endIndex)}${totalDepth}`;
        worksheet.mergeCells(startCell + ':' + endCell);
      }
    });
  });
}

/**
 * 数字转换为字母
 * @param columnNumber
 * @returns
 */
function convertToTitle(columnNumber) {
  const ans = [];

  while (columnNumber > 0) {
    const a0 = ((columnNumber - 1) % 26) + 1;

    ans.push(String.fromCharCode(a0 - 1 + 'A'.charCodeAt(0)));
    columnNumber = Math.floor((columnNumber - a0) / 26);
  }
  ans.reverse();
  return ans.join('');
}

// 获取每个单元格总共有多长
function getChildLength(children: any[], ans) {
  if (!children) return ans;
  children.forEach(child => {
    if (!child.children) {
      ans++;
    } else {
      ans += getChildLength(child.children, ans);
    }
  });
  return ans;
}

/**
 * 设置每个节点的父节点，用于startIndex的判断，同时获取跟节点
 * @param headerColumn
 * @param parent
 * @param leafNodes
 */
function setChildParent(
  headerColumn: TableColumnAllProps[],
  parent,
  leafNodes: TableColumnAllProps[]
) {
  for (let i = 0; i < headerColumn.length; i++) {
    const item = headerColumn[i];

    item.parent = parent || null;
    if (item.children) {
      setChildParent(item.children, item, leafNodes);
    } else {
      leafNodes.push(item);
    }
  }
}

// eslint-disable-next-line no-restricted-globals
self.onmessage = function (e) {
  const { data, type } = e.data || {};

  if (type === 'start') {
    // 开始进行导出任务
    exportExcel({ ...data });
  }
};

function postMessage(type, data) {
  // eslint-disable-next-line no-restricted-globals
  self.postMessage({
    type,
    data
  });
}
